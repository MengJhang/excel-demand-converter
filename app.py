from io import BytesIO
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from collections import OrderedDict
from datetime import datetime

st.set_page_config(page_title="Excel 需求轉換工具", page_icon="📊", layout="wide")

st.title("Excel 需求轉換工具")
st.write("上傳 old.xlsx，網站會自動轉換並產生可下載的 new.xlsx。")
st.markdown("""
**使用方式**
1. 上傳來源 Excel（old.xlsx）
2. 按下「開始轉換」
3. 下載轉換後的 new.xlsx
""")
st.warning("若檔案包含公司機密、個資或受管制資料，請先確認公司允許使用公開雲端服務。此工具不會將檔案寫入永久資料庫，但檔案會在網站執行期間暫存在記憶體中。")

# 轉檔規則（沿用原本已確認的 Python 程式）
ITEM_HEADER = "Request items"
PASTDUE_HEADER = "Pastdue"
MATERIAL_KEY_HEADERS = ["P/N", "Vendor P/N", "Plant"]
DEMAND_TYPES = ["plan requirement", "raw mtl so", "production order"]
FIRM_ORDER_TYPE = "firm order"
EXCLUDED_TYPES = {"subcontract po", "plan order"}

def norm(value):
    return str(value or "").strip().lower()

def number(value):
    if value is None or value == "":
        return 0
    if isinstance(value, (int, float)):
        return value
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0

def parse_date_header(value):
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if len(text) == 8 and text.isdigit():
        try:
            return datetime.strptime(text, "%Y%m%d")
        except ValueError:
            return None
    return None

def convert_workbook(uploaded_bytes):
    # -----------------------------
    # 1. 讀取 old.xlsx
    # -----------------------------
    old_wb = openpyxl.load_workbook(BytesIO(uploaded_bytes), data_only=True)
    old_ws = old_wb.active

    old_headers = {
        str(old_ws.cell(1, c).value).strip(): c
        for c in range(1, old_ws.max_column + 1)
        if old_ws.cell(1, c).value is not None
    }

    for required in [ITEM_HEADER, PASTDUE_HEADER, *MATERIAL_KEY_HEADERS]:
        if required not in old_headers:
            raise ValueError(f"old.xlsx 找不到必要欄位：{required}")

    item_col = old_headers[ITEM_HEADER]
    pastdue_col = old_headers[PASTDUE_HEADER]
    key_cols = [old_headers[name] for name in MATERIAL_KEY_HEADERS]

    # 自動找出週別欄位：欄名為 YYYYMMDD 的欄位
    week_columns = []
    for c in range(1, old_ws.max_column + 1):
        dt = parse_date_header(old_ws.cell(1, c).value)
        if dt:
            week_columns.append((c, dt))

    if not week_columns:
        raise ValueError("old.xlsx 找不到週別欄位，請確認日期欄名格式。")

    # -----------------------------
    # 2. 依材料分組舊檔資料
    # -----------------------------
    materials = OrderedDict()

    for r in range(2, old_ws.max_row + 1):
        # 忽略沒有材料編號的空白列
        if old_ws.cell(r, key_cols[0]).value is None:
            continue

        key = tuple(old_ws.cell(r, c).value for c in key_cols)
        item_type = norm(old_ws.cell(r, item_col).value)

        # Subcontract PO / Plan order 不會輸出，也不參與 Demand
        if item_type in EXCLUDED_TYPES:
            pass

        if key not in materials:
            materials[key] = {
                "base_row": r,
                "rows": {},
            }

        # 同一材料若有重複類型，保留第一列；正常資料通常每種類型一列
        materials[key]["rows"].setdefault(item_type, r)

    # -----------------------------
    # 3. 建立全新的輸出活頁簿
    # -----------------------------
    new_wb = openpyxl.Workbook()
    ws = new_wb.active
    ws.title = "Material Demand"

    base_headers = [
        "Drawing Version", "P/N", "Vendor P/N", "Vendor code",
        "Vendor name", "Description", "Plant", "Round value",
        "Plant stock", "Insption stock", "Firm Order", "PR Date",
        "缺料週", "Request items", "Pastdue"
    ]
    week_dates = [dt for _, dt in week_columns]
    meta_headers = [
        "Safety stock", "ABC Indicator", "MRP Controller",
        "MRP controller Name"
    ]

    all_headers = base_headers + week_dates + meta_headers
    for c, header in enumerate(all_headers, start=1):
        cell = ws.cell(1, c, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 日期欄標題格式
    first_week_col = len(base_headers) + 1
    last_week_col = first_week_col + len(week_dates) - 1
    for c in range(first_week_col, last_week_col + 1):
        ws.cell(1, c).number_format = "yyyy/mm/dd"

    # 欄位位置
    out_headers = {str(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)}
    out_item_col = out_headers["Request items"]
    out_pastdue_col = out_headers["Pastdue"]
    out_firm_col = out_headers["Firm Order"]
    out_prdate_col = out_headers["PR Date"]
    out_shortage_col = out_headers["缺料週"]

    # 需求與庫存基本欄位從 old.xlsx 直接對應
    copy_headers = [
        "Drawing Version", "P/N", "Vendor P/N", "Vendor code",
        "Vendor name", "Description", "Plant", "Round value",
        "Plant stock", "Insption stock", "Safety stock", "ABC Indicator",
        "MRP Controller", "MRP controller Name"
    ]

    # -----------------------------
    # 4. 每個材料輸出 Demand / Supply / Net avail 三列
    # -----------------------------
    output_row = 2
    thin_gray = Side(style="thin", color="D9E2F3")

    for key, material in materials.items():
        source_rows = material["rows"]
        base_r = material["base_row"]

        # 若該材料沒有任何需求/firm order 記錄，仍可建立三列
        row_demand = output_row
        row_supply = output_row + 1
        row_net = output_row + 2
        output_row += 3

        # 由材料的第一列複製基本資訊，PR Date 保持空白
        for name in copy_headers:
            if name in old_headers and name in out_headers:
                value = old_ws.cell(base_r, old_headers[name]).value
                for rr in (row_demand, row_supply, row_net):
                    ws.cell(rr, out_headers[name]).value = value

        for rr, label in [
            (row_demand, "Demand"),
            (row_supply, "Supply"),
            (row_net, "Net avail"),
        ]:
            ws.cell(rr, out_item_col).value = label
            # PR Date 是使用者手動備註欄，程式不填入任何內容
            ws.cell(rr, out_prdate_col).value = None

        # Firm Order = 該材料舊檔 Firm order 列的 Pastdue + 所有週別值加總
        firm_row = source_rows.get(FIRM_ORDER_TYPE)
        firm_total = 0
        if firm_row:
            firm_total += number(old_ws.cell(firm_row, pastdue_col).value)
            for old_c, _ in week_columns:
                firm_total += number(old_ws.cell(firm_row, old_c).value)

        for rr in (row_demand, row_supply, row_net):
            ws.cell(rr, out_firm_col).value = firm_total

        # Demand Pastdue：只加總三種指定需求的 Pastdue
        pastdue_demand = sum(
            number(old_ws.cell(source_rows[item], pastdue_col).value)
            for item in DEMAND_TYPES
            if item in source_rows
        )
        ws.cell(row_demand, out_pastdue_col).value = pastdue_demand

        # Supply 新增列，不填數值；Pastdue 同樣留白
        ws.cell(row_supply, out_pastdue_col).value = None

        # Net avail 的 Pastdue 公式：Plant stock + Inspection stock - Pastdue Demand + Pastdue Supply
        plant_stock_col = out_headers["Plant stock"]
        inspection_stock_col = out_headers["Insption stock"]
        plant_stock_letter = get_column_letter(plant_stock_col)
        inspection_stock_letter = get_column_letter(inspection_stock_col)
        pastdue_letter = get_column_letter(out_pastdue_col)
        ws.cell(row_net, out_pastdue_col).value = (
            f"={plant_stock_letter}{row_net}+{inspection_stock_letter}{row_net}"
            f"-{pastdue_letter}{row_demand}"
            f"+{pastdue_letter}{row_supply}"
        )

        # 每週 Demand = 同一週的 Plan requirement + Raw Mtl SO + Production order
        for i, (old_c, dt) in enumerate(week_columns):
            new_c = first_week_col + i
            demand_value = 0
            for item in DEMAND_TYPES:
                source_r = source_rows.get(item)
                if source_r:
                    demand_value += number(old_ws.cell(source_r, old_c).value)

            # 僅加總當週，不累計之前週別
            ws.cell(row_demand, new_c).value = demand_value
            # Supply 整列每週留白
            ws.cell(row_supply, new_c).value = None

            # Net avail = 前一期 Net avail - 本週 Demand + 本週 Supply
            prev_col = out_pastdue_col if i == 0 else new_c - 1
            prev_letter = get_column_letter(prev_col)
            week_letter = get_column_letter(new_c)
            ws.cell(row_net, new_c).value = (
                f"={prev_letter}{row_net}"
                f"-{week_letter}{row_demand}"
                f"+{week_letter}{row_supply}"
            )

        out_pastdue_col_letter = get_column_letter(out_pastdue_col)

        # 缺料週：從 Pastdue 開始，找第一個 Net avail < 0 的欄位；無缺料顯示 -
        shortage_start_letter = get_column_letter(out_pastdue_col)
        week_end_letter = get_column_letter(last_week_col)
        shortage_formula = (
            f'=IFERROR(INDEX(${shortage_start_letter}$1:${week_end_letter}$1,'
            f'MATCH(TRUE,{shortage_start_letter}{row_net}:'
            f'{week_end_letter}{row_net}<0,0)),"-")'
        )

        for rr in (row_demand, row_supply, row_net):
            ws.cell(rr, out_shortage_col).value = shortage_formula
            ws.cell(rr, out_shortage_col).number_format = "yyyy/mm/dd"

        # 套用基本格式及材料區塊分隔線
        separator_side = Side(style="medium", color="7F8C9A")
        for rr in (row_demand, row_supply, row_net):
            for c in range(1, len(all_headers) + 1):
                cell = ws.cell(rr, c)
                cell.border = Border(bottom=thin_gray)
                if c in range(first_week_col, last_week_col + 1):
                    cell.number_format = '#,##0;[Red]-#,##0;–'
            ws.cell(rr, out_prdate_col).number_format = "yyyy/mm/dd"

        # 每個材料的第三列（Net avail）底部加粗分隔線，方便區分材料區塊
        for c in range(1, len(all_headers) + 1):
            cell = ws.cell(row_net, c)
            cell.border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=separator_side,
            )

        ws.cell(row_net, out_item_col).font = Font(bold=True)
        ws.cell(row_demand, out_item_col).font = Font(bold=True)

    # -----------------------------
    # 5. 欄寬、凍結窗格、篩選及計算設定
    # -----------------------------
    widths = {
        "A": 14, "B": 16, "C": 30, "D": 16, "E": 20,
        "F": 42, "G": 12, "H": 12, "I": 14, "J": 15,
        "K": 14, "L": 18, "M": 14, "N": 16, "O": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for c in range(first_week_col, last_week_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 13
    for c in range(last_week_col + 1, len(all_headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18

    ws.freeze_panes = "P2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30

    # 讓 Excel 開啟檔案時重新計算公式
    try:
        new_wb.calculation.fullCalcOnLoad = True
        new_wb.calculation.forceFullCalc = True
        new_wb.calculation.calcMode = "auto"
    except AttributeError:
        pass

    output = BytesIO()
    new_wb.save(output)
    output.seek(0)
    return output.getvalue(), len(materials), output_row - 2

uploaded = st.file_uploader("選擇來源檔案（.xlsx）", type=["xlsx"])

if uploaded is not None:
    st.success(f"已選擇：{uploaded.name}")
    if st.button("開始轉換", type="primary"):
        try:
            with st.spinner("正在轉換檔案..."):
                output_bytes, material_count, row_count = convert_workbook(uploaded.getvalue())
            st.success(f"轉換完成！材料數：{material_count}；輸出列數：{row_count}")
            st.download_button(
                "下載轉換後的 new.xlsx",
                data=output_bytes,
                file_name="new.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        except Exception as e:
            st.error(f"轉換失敗：{e}")
            st.info("請確認上傳的是正確的 old.xlsx，且欄位名稱和週別格式符合原始程式的要求。")
else:
    st.caption("請先上傳 old.xlsx。")

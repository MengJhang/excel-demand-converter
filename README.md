# Excel 需求轉換網站

## 本機測試
```bash
pip install -r requirements.txt
streamlit run app.py
```
瀏覽器會開啟本機網站。上傳 `old.xlsx`，按「開始轉換」，再下載 `new.xlsx`。

## 部署成公開網站（Streamlit Community Cloud）
1. 先登入或建立 GitHub 帳號。
2. 在 GitHub 建立一個新的 repository，例如 `excel-demand-converter`。
3. 將本資料夾中的 `app.py` 與 `requirements.txt` 上傳到 repository 根目錄。
4. 開啟 https://share.streamlit.io/ 並登入。
5. 選擇建立 App，連結剛才的 GitHub repository。
6. Main file path 設為 `app.py`，按 Deploy。
7. 部署成功後，平台會提供一個公開網址，使用者即可上傳、轉換及下載。

## 注意事項
- 此網站不會主動把上傳檔案保存到永久資料庫，但檔案會在執行期間於伺服器記憶體處理。
- 如果 Excel 包含公司機密或個人資料，請先確認公司資安政策是否允許上傳至公開雲端。
- 公開網址代表任何取得網址的人都可能使用此工具；如需限制特定使用者，應再加入登入驗證或部署至公司核准的環境。

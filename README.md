# Protection Substation Streamlit App

此專案已整理為可部署到 GitHub + Streamlit 的版本。

## 檔案說明

- `app.py`：Streamlit 主程式（上傳來源檔案、處理、下載結果）
- `requirements.txt`：部署需要的 Python 套件
- `.streamlit/config.toml`：Streamlit 設定檔
- `.gitignore`：Git 忽略清單
- `assets/`：放提供使用者下載的公版檔案

## 公版檔案下載設定

若要在 App 顯示下載按鈕，請將檔案放在 `assets/` 並使用以下檔名：

- `assets/公版_IED試驗報告.xlsm`
- `assets/公版_變電所_測試程序.psx`

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 使用方式

1. 上傳「來源資料檔案」（可多選）
2. 上傳「目標檔案」（建議 `.xlsm` 或 `.xlsx`）
3. 點選「開始轉換」
4. 下載產生的 `updated_*.xlsx` / `updated_*.xlsm` 檔案

## 部署到 Streamlit Cloud

1. 將此資料夾推到 GitHub Repository
2. 到 Streamlit Cloud 建立新 App
3. 設定：
   - **Repository**：你的 repo
   - **Branch**：main（或你使用的分支）
   - **Main file path**：`app.py`
4. Deploy

## 注意事項

- 目前目標檔案僅支援 `.xlsx` / `.xlsm`（`openpyxl` 限制）
- 若 `O1` 是公式，需先在 Excel 儲存過一次，讓快取值可被讀取
- Streamlit 雲端環境不支援 `xlwings` / `tkinter` 視窗操作，因此改為網頁上傳下載流程


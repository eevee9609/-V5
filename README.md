# 團檢模板自動生成器

`health-check-template-generator` 是協助團體健康檢查作業的 Windows 桌面工具，提供檢查項目與院所設定、公費／自費模板產生、名冊處理及杏聯／博仁拆分功能。

原始程式由 `團檢整合完美ver5.1.py` 重構為模組化專案；儲存庫名稱以用途命名，後續版本透過 Git 提交紀錄追蹤。

## 環境與安裝

- Windows、Python 3.10 以上（需包含 Tcl/Tk）。
- 使用 xlwings 的 Excel 自動化功能時，需安裝桌面版 Microsoft Excel。
- 建議安裝 MiSans 字型，以符合原始介面設定。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 啟動

在專案根目錄執行：

```powershell
.\.venv\Scripts\python.exe main.py
```

## 原則

- `app/config`：路徑、外觀與 Tcl/Tk 初始化。
- `app/data`：JSON 資料庫讀寫與項目查詢。
- `app/utils`：不依賴畫面的通用函式。
- `app/ui`：各個視窗與元件。
- `app/services`：Excel 產生與其他商業流程。
- `data`、`assets`：執行時需要的可編輯資料與靜態資源。

原始程式使用的 JSON、圖示與熊貓圖片均保留在這個專案內。`data` 是執行時可編輯的設定，更新前請先備份自己調整過的內容。

## 檔案範圍

儲存庫包含程式碼、JSON 設定／專案預設、圖片資源、打包設定與 `qa` 驗證腳本。依上傳設定，Excel 模板、名冊、輸出報表（`.xls`、`.xlsx`、`.xlsm`、`.xlsb`）、EXE、壓縮檔及自動備份不納入版本控制。需要的 Excel 檔請在本機自行選取或準備。

## 驗證

以下回歸測試使用合成資料，不會修改實際業務 Excel 檔案：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s qa -p "test_*.py" -v
```

`qa/check_settings_layout.py` 是需要桌面與 Tcl/Tk 的介面檢查；`qa/check_special_formula_results.mjs` 是選用的獨立公式驗證，另需 Node.js 與提供 `@oai/artifact-tool` 的環境。

## 打包

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean 團檢模板自動生成器.spec
```

此版本採用單一 EXE；請將 `data` 與 `assets` 資料夾放在產生的 EXE 同一層。這能讓院所資料、系統設定與檢查項目資料庫在更新程式時仍可保留，且應用程式可寫入這些 JSON 檔。

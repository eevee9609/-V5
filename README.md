# 團檢模板自動生成器（結構化版本）

這個資料夾是從 `團檢整合完美ver5.1.py` 逐步重構而來的版本；原始單檔保持不變，方便隨時回復與比對。

## 啟動

在已安裝原有套件的 Python 環境中執行：

```powershell
python main.py
```

## 原則

- `app/config`：路徑、外觀與 Tcl/Tk 初始化。
- `app/data`：JSON 資料庫讀寫與項目查詢。
- `app/utils`：不依賴畫面的通用函式。
- `app/ui`：各個視窗與元件。
- `app/services`：Excel 產生與其他商業流程。
- `data`、`assets`：執行時需要的可編輯資料與靜態資源。

原始程式使用的 JSON、圖示與熊貓圖片均會保留一份在這個專案內。重構期間不要直接刪除舊版檔案。

## 打包

```powershell
pyinstaller --noconfirm --clean 團檢模板自動生成器.spec
```

此版本採用單一 EXE；請將 `data` 與 `assets` 資料夾放在產生的 EXE 同一層。這能讓院所資料、系統設定與檢查項目資料庫在更新程式時仍可保留，且應用程式可寫入這些 JSON 檔。

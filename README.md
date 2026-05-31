# ropin13.github.io

本專案是可直接部署於 GitHub Pages 的多頁式靜態網站，首頁為 `index.html`，整合資料分析工具、健康追蹤與小遊戲入口。

## 功能總覽

### 1. 功能首頁
- 檔案：`index.html`
- 用途：集中導覽所有功能頁。
- 特色：Bootstrap 卡片式入口，含資料比對、繳費通知比對、夢幻水族館、健康數據頁面連結。

### 2. 資料比對分析工具
- 檔案：`vueindex_new.html`、`app_new.js`
- 用途：上傳 Excel 錯誤分析報表，依錯誤欄位比對主機值與再造值。
- 主要功能：
  - 讀取 `Summary` 與 `Data_XXXXXXXXX` 工作表。
  - 動態欄位組（內建多組、可自訂）。
  - 相似度（Levenshtein）計算與條件過濾（例如 `>=90`）。
  - `unknown` / `miss` 筆數統計與明細檢視。
  - 分頁、欄位過濾、明細 Modal 檢視。

### 3. 繳費通知比對頁
- 檔案：`vueindex_payment_notice.html`、`app_pn.js`
- 用途：上傳 Excel 後針對繳費通知欄位進行篩選與查詢。
- 主要功能：
  - 錯誤欄位切換。
  - 多欄位過濾（包含 `CTRL_CODE` 等欄位）。
  - 表格分頁與明細檢視。

### 4. 健康數據追蹤面板
- 檔案：`health.html`
- 用途：記錄血壓、脈搏、血糖與備註，並查看趨勢圖。
- 主要功能：
  - 支援本地暫存資料與 Google Apps Script API 模式。
  - Chart.js 趨勢圖（最近七天、14 天、1 個月、1 年）。
  - 歷史清單排序與分頁。
  - 最近七天平均血壓統計。

### 5. 夢幻水族館遊戲
- 檔案：`aquarium-game.html`
- 用途：純前端放置型養成小遊戲。
- 主要功能：
  - 商店買魚、成長、出售與金幣循環。
  - 升級、圖鑑、成就與主題客製。
  - `localStorage` 存檔與離線成長（需購買自動成長升級）。

## 技術與相依套件

本專案無打包流程，主要以 CDN 載入前端函式庫：

- Vue 3
- Bootstrap 5
- Bootstrap Icons
- SheetJS (xlsx)
- Tailwind CSS（健康頁）
- Chart.js（健康頁）
- Lucide Icons（健康頁）

## 本機執行

使用任一靜態伺服器即可：

```bash
python3 -m http.server 8000
```

開啟：

- `http://127.0.0.1:8000/index.html`

## 使用注意事項

### 資料比對工具 / 繳費通知比對頁
- 上傳的 Excel 檔案需先移除密碼保護。
- 需符合工具預期格式：`Summary` 與多個 `Data_XXXXXXXXX` 工作表。

### 健康數據頁
- 若未設定 API URL，系統會使用本地示範/暫存資料模式。
- API URL 可在頁面右上角設定，並儲存於瀏覽器 `localStorage`。

## 主要檔案說明

- `index.html`：功能首頁
- `vueindex_new.html`：資料比對分析工具 UI
- `app_new.js`：資料比對分析工具邏輯
- `vueindex_payment_notice.html`：繳費通知比對頁 UI
- `app_pn.js`：繳費通知比對頁邏輯
- `health.html`：健康數據追蹤頁（含圖表、API 設定、表單）
- `aquarium-game.html`：夢幻水族館遊戲
- `vueindex_new.md`：資料比對分析工具技術文件

## 部署

可直接部署至 GitHub Pages（Repository Root）。
`index.html` 會作為預設首頁入口。

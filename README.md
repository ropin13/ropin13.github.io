# ropin13.github.io

這是一個以 GitHub Pages 可直接部署的靜態前端專案，已將主要功能拆分為獨立頁面，並以 `index.html` 作為統一首頁入口。

## 首頁與功能分流

- **首頁**：`index.html`
  - 作為預設入口，提供各功能頁連結卡片。
- **資料比對分析工具**：`vueindex_new.html` + `app_new.js`
  - 讀取 Excel 錯誤報表，進行欄位比對、相似度分析、條件過濾與明細檢視。
- **繳費通知比對頁**：`vueindex_payment_notice.html` + `app_pn.js`
  - 針對繳費通知場景做過濾、統計與查詢。
- **夢幻水族館遊戲**：`aquarium-game.html`
  - 純前端放置型小遊戲。

## 本機啟動方式

由於是靜態網站，可直接使用任何 HTTP Server 啟動：

```bash
python3 -m http.server 8000
```

啟動後開啟：

- `http://127.0.0.1:8000/index.html`

## 專案檔案說明

- `index.html`：首頁（功能導覽）
- `vueindex_new.html`：資料比對工具 UI
- `app_new.js`：資料比對工具邏輯
- `vueindex_payment_notice.html`：繳費通知工具 UI
- `app_pn.js`：繳費通知工具邏輯
- `aquarium-game.html`：水族館遊戲頁
- `vueindex_new.md`：資料比對工具技術文件

## 部署

本專案可直接部署到 GitHub Pages。若 repository 設定以 root 提供靜態頁面，`index.html` 會自動成為首頁。

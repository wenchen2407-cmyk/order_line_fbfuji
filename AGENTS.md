# Project LINE Order - AI 開發與升級強制執行規範 (Project Rules)

本工作區（Workspace）是基於 Google Apps Script (GAS)、Google 試算表 (Google Sheets)、LINE LIFF 與前端響應式網頁所架構的電商與代購管理系統。

為避免出現「介面更新頻繁，卻未與後端資料庫（Google Sheets / GAS）實質串接整合」的嚴重脫節問題，**任何在此專案進行操作的 AI 助理，必須嚴格遵守以下標準作業規範**：

---

## 🛑 強制執行閘門（AI 實作前必先詢問確認）

在未經使用者明確回答前，**嚴禁**逕自宣告「功能已完成」或隨便生成帶有死假資料（Mock Data）的孤立檔案。
遇到功能開發或介面更新需求時，AI 第一步必須主動詢問使用者：

1. **升級類型確認**：
   - 這是「**純視覺雛形展示（Preview Only）**」？還是「**正式功能上線整合（Production Upgrade）**」？
2. **資料欄位確認**：
   - 本次更動是否有**新增、刪除或修改任何後端資料欄位**（例如訂單狀態、商品成本、物流單號、備註標籤等）？
3. **目標實體檔案**：
   - 本次更動應該落實在哪一個正式前端檔案（例如 `admin-card-generator.html`、`my-orders.html`、`index.html`）？
4. **既有資料相容性**：
   - Google Sheets 中現有的歷史資料是否需要自動補齊預設值？

---

## 📋 開發與雙軌部署六步升級法 (6-Step Upgrade & Deployment Protocol)

### 步驟 1：建立前後端欄位對照字典 (Schema Alignment)
若涉及新欄位，必須在程式碼中與回覆中清楚列出：
- **Google Sheets 中文表頭名稱**（例如：`採購/出貨狀態`）
- **前端 JavaScript 物件 Key**（例如：`purchaseStatus`）
- **資料格式與預設值**（例如：字串、數字、或需標記 `@` 純文字防開頭 0 遺失）

### 步驟 2：後端模組化維護與自動編譯打包 (Modular Backend & Build)
- **嚴禁直接手動修改整包 `src/gas/Code.gs`**，必須至 `src/gas/modules/` 修改對應的子模組：
  - `00_Config.gs`：系統常數、白名單與共用工具。
  - `01_Router.gs`：Web App 路由 (`doGet` / `doPost`)。
  - `02_ProductService.gs`：商品管理（CRUD、多圖、原幣成本公式）。
  - `03_OrderService.gs`：訂單保留、加單合併、結帳出貨、採購作業站。
  - `04_CustomerService.gs`：顧客歸戶、會員動態 Log、1對1結帳通報名單。
  - `05_TelegramService.gs`：Telegram 即時推播、HTML 排版、訪客稱呼比對。
  - `06_SystemAdminService.gs`：試算表結構 (`setupSpreadsheet`)、無損擴充與系統設定。
- 確保包含無損自動擴充欄位的邏輯（`insertColumnsAfter`），讀取資料以「表頭名稱動態查找索引（`indexOf`）」，嚴禁寫死數字索引。
- 修改完成後，**必須執行自動編譯檢驗腳本**：
  ```bash
  node scripts/build-gas.js
  ```
  確保 100% 通過 V8 / Node.js 語法檢測，並自動生成最新的 `src/gas/Code.gs`。

### 步驟 3：前端正式程式碼整合
- 將新介面功能合併到指定的正式前端檔案中。
- **嚴禁在正式功能中保留寫死的 `let mockOrders = [...]` 假資料**，必須正式接上 `fetch(GAS_API_URL)`。
- 實作「載入中（Loading）」、「錯誤提示」與「操作成功回饋（Toast）」。

### 步驟 4：測試機先行部署與自動化驗測閘門 (Test-First Gate)
- **嚴禁在未經測試機驗測前直接部署至正式機！**
- 任何涉及後端與核心前端之更動，**第一優先必須先在【測試專用環境】完成部署與驗測**：
  1. 將編譯後程式碼部署至【測試專用 Google Apps Script】（綁定獨立之【測試專用試算表】）。
  2. 執行自動化測試腳本（例如 `node scripts/test-high-frequency-logs.js` 等 API / 併發測試）。
  3. 執行電腦端（1440px）與手機端（375px / 390px，防橫向溢出、100dvh、44px 熱區）響應式與無報錯自檢。
  4. **出具「測試機自動化驗測成績單」**：必須 100% 標註通過（PASS），方可進入正式機發布。

### 步驟 5：正式機部署發布指引 (測試通過後放行)
當測試機自動化驗測 100% PASS 後，**AI 必須明確提供標準兩步驟正式部署指引**，引導使用者將驗測無誤的代碼同步至正式環境：
> **🚀 正式機 (Production) 部署操作指引（測試通過後執行）：**
> 1. **複製最新通過驗測之程式碼**：
>    - 打開本機 `src/gas/Code.gs`，按 `Ctrl + A` (`Cmd + A`) 全選複製。
>    - 切換至【正式營運 Google Apps Script 編輯器】，點選 `程式碼.gs`，全選貼上覆蓋並儲存 (`Ctrl + S` / `Cmd + S`)。
> 2. **發布 Web App「新版本」**：
>    - 點擊右上角藍色按鈕「部署」➔「管理部署作業」。
>    - 點擊右側 ✏️ 鉛筆圖示（編輯）。
>    - 在「版本」下拉選單中，務必選擇「**新版本**」 (New version)。
>    - 點擊右下角「**部署**」完成正式機上線！
> 3. **前端發布**：將變更推送到 GitHub main 分支以更新 GitHub Pages。

### 步驟 6：正式環境生產線上巡檢 (Production Smoke Test)
完成正式發布後，進行輕量線上巡檢：
- 檢查正式環境 API 呼叫是否順暢、Google 試算表欄位是否正確讀取回顯，確保 100% 平滑無痛切換。

# 系統部署與設定教學手冊 (Step-by-Step Guide)

本手冊將手把手引導您在 **15 分鐘內** 完成整套系統的串接與上線。

---

## 📌 步驟 1：建立 Google 試算表與部署 GAS 後端

1. 打開 [Google 雲端硬碟](https://drive.google.com/)，新增一個空白的 **Google 試算表**（例如命名為：`連線代購資料庫`）。
2. 點擊頂部選單的 **「擴充功能」 > 「Apps Script」**。
3. 將預設編輯器裡的程式碼清空，將本專案 [`src/gas/Code.gs`](file:///c:/Users/WW_love_TC/Documents/Project%20LINE%20Order/src/gas/Code.gs) 的所有內容完整複製並貼入。
4. 儲存專案（命名為 `LINE-Order-API`）。
5. **初始化表格**：
   * 在上方工具列的函式下拉選單中選擇 **`setupSpreadsheet`**。
   * 點擊 **「執行」**。
   * （首次執行會跳出權限授權提示，點「檢查權限」> 選擇您的 Google 帳號 > 點「進階」>「前往專案」並允許）。
   * 執行完畢後切回試算表，您會看到「商品清單」、「訂單明細」、「顧客歸戶」、「系統設定」這 4 個工作表已自動建立完成並帶入預設範例！
6. **部署為 Web 應用程式**：
   * 點擊右上角的 **「部署」 > 「新部署」**。
   * 齒輪圖示選擇 **「網路應用程式」 (Web App)**。
   * **說明**：填寫 `v1`。
   * **執行身分**：選擇 **「我」 (Me)**。
   * **誰可以存取**：務必選擇 **「所有人」 (Anyone)**。
   * 點擊「部署」，複製生成的 **網路應用程式網址**（通常以 `https://script.google.com/macros/s/.../exec` 結尾）。

---

## 📌 步驟 2：申請 LINE Developers 與建立 LIFF

1. 前往 [LINE Developers Console](https://developers.line.biz/) 並登入您的 LINE 帳號。
2. 點擊 **Create a new Provider**（或選擇既有的 Provider）。
3. 建立一個 **LINE Login** Channel（通道）：
   * Channel name: 您的連線商店名稱
   * App type: 勾選 Web app
4. 進入建立好的 Channel，切換到 **LIFF** 分頁，點擊 **Add** 新增 LIFF 應用程式：
   * **LIFF app name**：`連線下單`
   * **Size**：建議選擇 **Full** 或 **Tall**
   * **Endpoint URL**：填寫您放置前端頁面的網址（例如 GitHub Pages、Firebase Hosting 或 Netlify 網址；本機測試時可先填寫您的測試伺服器網址）
   * **Scopes**：勾選 `profile`、`openid`
   * ⭐ **關鍵必開設定**：尋找 **`Share Target Picker`** 選項，將其設定為 **Enabled (開啟)**（此功能是讓賣家能將 Flex 卡片免推播費發送到群組的核心！）
5. 建立完成後，複製 **LIFF ID**（格式如：`1234567890-AbCdEfGh`）。

---

## 📌 步驟 3：前端配置與上線

開啟前端檔案進行設定值替換：
1. 開啟 [`src/frontend/index.html`](file:///c:/Users/WW_love_TC/Documents/Project%20LINE%20Order/src/frontend/index.html)：
   * 將 `YOUR_LIFF_ID_HERE` 替換為您的 LIFF ID。
   * 將 `YOUR_GAS_EXEC_URL_HERE` 替換為步驟 1 的 Web App 網址。
2. 開啟 [`src/frontend/my-orders.html`](file:///c:/Users/WW_love_TC/Documents/Project%20LINE%20Order/src/frontend/my-orders.html) 同樣替換這兩項參數。
3. 開啟 [`src/frontend/admin-card-generator.html`](file:///c:/Users/WW_love_TC/Documents/Project%20LINE%20Order/src/frontend/admin-card-generator.html) 替換 LIFF ID。

---

## 🚀 賣家平日營運操作流程（極簡工作流）

1. **上架好物**：
   * 在 Google 試算表的「商品清單」輸入商品名稱、特價、庫存與圖片。
2. **產出商品卡並發群**：
   * 用手機開啟 `admin-card-generator.html`，輸入商品編號預覽。
   * 點擊 **「🚀 一鍵分享商品卡到 LINE 群組」**，直接勾選群組發出！
3. **客人下單**：
   * 客人在群組看到漂亮的 Flex 互動卡，點「立即下單」，自動免登入帶入 LINE 暱稱完成下單。
4. **即時對帳與出貨**：
   * 試算表「訂單明細」會即時更新。
   * 客人匯款後至「我的訂單」填寫後五碼，試算表狀態自動切換為「已回報待對帳」，核對存摺入帳後賣家手動改為「已付款」，準備出貨！

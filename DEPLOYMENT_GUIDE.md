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

1. **商品快速建檔 (電腦端操作)**：
   * 開啟 `admin-card-generator.html` 進入「💻 電腦快速建檔」分頁。
   * 輸入商品名稱、連線代購價（若無專櫃原價直接留空即可）。
   * 可利用內建的 **「🧮 連線代購成本計算機」**（輸入日幣/韓元原幣金額與重量，自動以 240元/kg 空運算出總成本與毛利，自動記至試算表）。
   * 勾選「連線時間不限名額跟庫存」或設定庫存。
   * 點選「📁 選擇照片上傳」支援多選照片，系統會自動在瀏覽器智慧壓縮（100~200KB）並秒速上傳至指定 Google Drive 資料夾。
   * 點擊儲存，立即同步寫入 Google 試算表！

2. **多品打包輪播發群 (手機端操作)**：
   * 在手機 LINE 聊天室中點開 `admin-card-generator.html`。
   * 勾選 1~10 件熱門商品，右側可預覽輪播效果。
   * 點擊 **「🚀 一鍵打包推播至 LINE 群組」**，直接選擇指定群組送出高質感左右滑動輪播大卡（Carousel）！

3. **客人免登入極速搶購 (連線下單體驗)**：
   * 客人在群組點擊卡片上的「🛒 立即下單」直接進入 LIFF 下單頁（`index.html`）。
   * 圖片支援等比完整無損呈現（Contain 模式不裁切），多圖提供縮圖相簿列切換。
   * **連線登記 0 秒無痛**：連線期間搶購無需填寫配送地址與運費！
   * **新客 1 次建檔**：若是第一次購買，點擊下單會彈出「首次建立登記資料」視窗（自動讀取 LINE 暱稱，僅需輸入真實姓名與手機一次並自動記憶）；老客戶則享有 0.5 秒光速搶購體驗！

4. **回國合併結帳與 6 大配送方式 (`my-orders.html`)**：
   * 連線結束回國後，客人開啟「我的訂單 / 出貨結帳清單」頁面，系統自動彙整該客人的所有搶購商品、款式、數量與小計。
   * **6 大配送方式動態試算**：
     1. 超商純取貨 (需先匯款)：運費 0 元 (不限金額)
     2. 7-11超商貨到付款 (賣貨便)：運費 38 元 / 滿 3,000 元免運
     3. 全家超商貨到付款 (好賣+)：運費 35 元 / 滿 3,000 元免運
     4. 郵局純寄件 (需先匯款)：郵資 80 元
     5. 郵局貨到付款：郵資 130 元
     6. 黑貓宅配到府 (需先匯款)：運費 100 元
   * **智慧收件資料切換**：選擇超商提供門市代號/名稱填寫與「7-11/全家門市查詢」捷徑；選擇郵局/宅配則提供縣市與詳細地址填寫。
   * **付款與賣場說明**：
     * **先匯款方式**：展開深色金屬感銀行帳戶卡片（支援一鍵複製帳號）與即時匯款後五碼回報。
     * **賣貨便 / 好賣+ 貨到付款**：顯示說明「整理完畢後，賣家將於群組或私訊發送專屬賣貨便/好賣+賣場連結供您下單出貨！」，並標註取貨付款總額。
     * **郵局貨到付款**：提示郵差送達代收現金。
   * **單筆運費合併**：一次結帳僅計一次運費，寫入試算表第一筆訂單，其餘運費計 0，對帳與財務統計完全清晰！

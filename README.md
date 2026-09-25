# LINE 群組團購與連線代購下單系統 (LINE Order Suite)

一套專為 LINE 連線代購、群組團購小賣家量身打造的**零伺服器維護成本、免官方推播費用**之整合下單系統。

---

## 🎯 核心亮點

1. **零成本推播（ShareTargetPicker）**：利用賣家個人 LINE 帳號直接將 Flex Message 商品卡分享到任何指定群組，**不消耗官方帳號每月的付費推播則數**。
2. **LINE 免登入下單（LIFF）**：客人點擊商品卡開啟下單頁面，自動讀取 LINE 暱稱與唯一識別碼（User ID），徹底解決客人暱稱改動無法對帳的痛點。
3. **Google 試算表即資料庫**：
   - 商品上架、庫存扣減、訂單狀態管理，賣家全在熟悉的 Google 試算表處理，不需安裝複雜的後台。
   - 支援 GAS (Google Apps Script) 後端 RESTful API。
4. **代購併單與核帳模組**：
   - 支援同一連線檔期多筆訂單自動歸戶與合併計算運費。
   - 客人專屬查單頁，支援後五碼回填與上傳截圖，自動標記待對帳。

---

## 📂 系統規劃與文檔

* 詳細架構與業務分析請參閱 Artifact 規劃書：[SYSTEM_DESIGN.md](file:///C:/Users/WW_love_TC/.gemini/antigravity-ide/brain/5a233964-ca81-4986-90e6-e668f780b50b/SYSTEM_DESIGN.md)

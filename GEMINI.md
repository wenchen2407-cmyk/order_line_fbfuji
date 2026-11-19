# Project LINE Order - AI 開發與升級強制執行規範

請參閱同目錄下的 [AGENTS.md](file:///Users/wen/projects/Project%20LINE%20Order/AGENTS.md)。
所有進入本工作區的 AI 助理必須嚴格遵守：
1. **實作前確認**：主動提問確認「純視覺雛形 vs 正式升級整合」、「目標實體檔案」、「欄位增減與資料相容性」。
2. **後端模組化與自動編譯**：後端修改必須至 `src/gas/modules/` 對應子模組維護，並執行 `node scripts/build-gas.js` 自動語法檢驗與產出 `Code.gs`。
3. **每次必備部署指引**：任何更動涉及後端時，回覆中必須完整提供「①複製 Code.gs 貼上覆蓋 ➔ ②管理部署作業選擇新版本部署」的標準兩步驟指引。
4. **拒絕死假資料**：進行正式升級時，嚴禁僅產出帶有 Mock Data 的預覽檔，必須落實前後端閉環整合與實測驗收。

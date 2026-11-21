# Project LINE Order - AI 開發與升級強制執行規範

請參閱同目錄下的 [AGENTS.md](file:///Users/wen/projects/Project%20LINE%20Order/AGENTS.md)。
所有進入本工作區的 AI 助理必須嚴格遵守：
1. **實作前確認**：主動提問確認「純視覺雛形 vs 正式升級整合」、「目標實體檔案」、「欄位增減與資料相容性」。
2. **後端模組化與自動編譯**：後端修改必須至 `src/gas/modules/` 對應子模組維護，並執行 `node scripts/build-gas.js` 自動語法檢驗與產出 `Code.gs`。
3. **全自動測試機先行驗測閘門 (Automated Test Gate)**：AI 必須自動執行 `build-gas.js`、`clasp push` 測試機、執行自動化測試腳本與雙端響應式檢測，出具 100% PASS 成績單。
4. **全自動正式機部署與發布**：測試通過後，AI 自動切換並執行 `clasp push` 正式機、`clasp deploy` 發布新版本，並自動提交推送 GitHub Pages，使用者完全無需手動操作終端機或複製貼上。
5. **拒絕死假資料**：進行正式升級時，嚴禁僅產出帶有 Mock Data 的預覽檔，必須落實前後端閉環整合與實測驗收。

# Project LINE Order - AI 開發與升級強制執行規範

請參閱同目錄下的 [AGENTS.md](file:///Users/wen/projects/Project%20LINE%20Order/AGENTS.md)。
所有進入本工作區的 AI 助理必須嚴格遵守：
1. 實作前主動提問確認「純視覺雛形 vs 正式升級整合」、「目標實體檔案」、「欄位增減與資料相容性」。
2. 進行正式升級時，嚴禁僅產出帶有 Mock Data 的預覽檔，必須落實「試算表表頭升級 (setupSpreadsheet) + Code.gs API + 前端正式介面整合 + GAS 新版本發布指引」的完整閉環。

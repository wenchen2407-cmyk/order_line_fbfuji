/**
 * ==============================================================================
 * Project LINE Order - Google Apps Script (GAS) 自動打包編譯腳本
 * ==============================================================================
 * 功能說明：
 * 1. 自動依序讀取 src/gas/modules/ 底下的所有模組檔案 (00_Config.gs, 01_Router.gs, ...)
 * 2. 合併為單一正式發布檔 src/gas/Code.gs
 * 3. 自動執行 Node.js 語法檢驗 (Syntax Check)，確保 100% 零語法錯誤後才產出
 * 
 * 執行方式：
 *   node scripts/build-gas.js
 * ==============================================================================
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const MODULES_DIR = path.join(__dirname, '..', 'src', 'gas', 'modules');
const OUTPUT_FILE = path.join(__dirname, '..', 'src', 'gas', 'Code.gs');

function buildGAS() {
  console.log('🚀 開始進行 Google Apps Script 模組化打包編譯...\n');

  if (!fs.existsSync(MODULES_DIR)) {
    console.error(`❌ 錯誤：找不到模組目錄 ${MODULES_DIR}`);
    process.exit(1);
  }

  // 依檔名排序讀取所有 .gs 檔案
  const files = fs.readdirSync(MODULES_DIR)
    .filter(file => file.endsWith('.gs'))
    .sort();

  if (files.length === 0) {
    console.error(`❌ 錯誤：${MODULES_DIR} 中沒有找到任何 .gs 檔案`);
    process.exit(1);
  }

  console.log(`📦 發現 ${files.length} 個模組檔案：`);
  files.forEach(f => console.log(`   • ${f}`));

  const buildTime = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  let combinedContent = `/**
 * ==============================================================================
 * LINE 連線代購 & 群組團購整合系統 - GAS 後端 API (正式整合部署檔)
 * ==============================================================================
 * ⚠️ 本檔案為【自動編譯生成檔】，請勿直接手動編輯此檔案！
 * 如需修改功能，請至 src/gas/modules/ 修改對應的子模組，並執行：
 *   node scripts/build-gas.js
 * 
 * 打包時間：${buildTime} (Asia/Taipei)
 * 包含模組：${files.join(', ')}
 * ==============================================================================
 */\n\n`;

  let totalFunctions = 0;

  files.forEach(file => {
    const filePath = path.join(MODULES_DIR, file);
    const content = fs.readFileSync(filePath, 'utf-8');
    
    // 計算函式數量
    const funcMatches = content.match(/function\s+([a-zA-Z0-9_$]+)\s*\(/g) || [];
    totalFunctions += funcMatches.length;

    combinedContent += `\n/* ==================== 模組起點：${file} ==================== */\n`;
    combinedContent += content.trim();
    combinedContent += `\n/* ==================== 模組結束：${file} ==================== */\n\n`;
  });

  // 寫入暫存檔以進行語法驗證 (Node.js 要求 .js 副檔名)
  const tempFile = path.join(__dirname, '..', 'src', 'gas', '.Code.tmp.js');
  fs.writeFileSync(tempFile, combinedContent, 'utf-8');

  try {
    // 執行 Node 語法檢驗
    execSync(`node --check "${tempFile}"`, { stdio: 'pipe' });
    // 驗證通過，正式寫入 Code.gs
    fs.renameSync(tempFile, OUTPUT_FILE);
    
    const stats = fs.statSync(OUTPUT_FILE);
    const lineCount = combinedContent.split('\n').length;
    const sizeKB = (stats.size / 1024).toFixed(1);

    console.log('\n==================================================');
    console.log(`✅ 打包編譯成功！`);
    console.log(`📄 輸出目標：src/gas/Code.gs`);
    console.log(`📊 統計資訊：共 ${files.length} 個模組、${totalFunctions} 個函式、${lineCount} 行代碼 (${sizeKB} KB)`);
    console.log(`🛡️ 語法檢驗：100% 通過 V8 / Node.js 語法檢測！`);
    console.log('==================================================\n');
  } catch (err) {
    if (fs.existsSync(tempFile)) {
      fs.unlinkSync(tempFile);
    }
    console.error('❌ 打包失敗：合併代碼未通過語法檢測！');
    console.error(err.message);
    if (err.stderr) {
      console.error(err.stderr.toString());
    }
    process.exit(1);
  }
}

buildGAS();

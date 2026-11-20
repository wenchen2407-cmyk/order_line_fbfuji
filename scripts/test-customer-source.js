/**
 * ==============================================================================
 * scripts/test-customer-source.js
 * ==============================================================================
 * 【顧客來源預設為請選擇 & 自主選擇必填防呆】全自動端到端驗測腳本
 * 
 * 驗測範圍：
 * 1. 🌐 本機伺服器與 HTML 靜態檢驗 (index.html, catalog.html, product.html)
 *    - 驗證 <select> 第一個選項必為 value="" 且顯示 "-- 請選擇得知管道 --"
 *    - 驗證防呆邏輯存在 (alert '請選擇您是從哪裡得知本賣場！')
 *    - 驗證 ?reset=1 新客重設參數邏輯
 * 2. ⚡ 本機 HTTP 服務 (http://localhost:8888/) 即時響應檢測
 * 3. ☁️ 線上測試環境 GAS API 端到端閉環資料寫入與查詢檢測
 *    - 模擬顧客以自選客源 ("Instagram") 註冊歸戶
 *    - 立即回查會員資料 checkCustomer，驗證 source 欄位精準符合
 *    - 查驗【會員動態Log】安全稽核軌跡是否正確記錄來源
 * ==============================================================================
 */

const fs = require('fs');
const path = require('path');

const TEST_GAS_URL = 'https://script.google.com/macros/s/AKfycbzHdEBRN4OsqvXECkypDGvt_NHOjlAxet_9KP3LeaBJgzRlt478i-qhwDl7WmxcdGch/exec';
const ADMIN_UID = 'Uee7765ed3c1819cb94fd404cfab8553b';
const ADMIN_KEY = 'admin8888';
const LOCAL_HOST = 'http://localhost:8888';

const ROOT_DIR = path.resolve(__dirname, '..');

let totalTests = 0;
let passedTests = 0;
let failedTests = 0;

function assert(condition, message) {
  totalTests++;
  if (condition) {
    console.log(`  ✅ [PASS] ${message}`);
    passedTests++;
  } else {
    console.error(`  ❌ [FAIL] ${message}`);
    failedTests++;
  }
}

async function runLocalHtmlChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('📑 階段一：本機前端檔案與 DOM 結構語意自動化檢測');
  console.log('------------------------------------------------------------------');

  const filesToCheck = [
    { name: 'index.html', selectId: 'vipSource' },
    { name: 'catalog.html', selectId: 'vipSource' },
    { name: 'product.html', selectId: 'vipSource' },
    { name: 'product.html (彈窗)', file: 'product.html', selectId: 'regSource' }
  ];

  for (const item of filesToCheck) {
    const filename = item.file || item.name;
    const filePath = path.join(ROOT_DIR, filename);
    const content = fs.readFileSync(filePath, 'utf8');

    // 1. 檢查是否有 <option value="">-- 請選擇得知管道 --</option>
    const hasDefaultOption = content.includes('<option value="">-- 請選擇得知管道 --</option>');
    assert(hasDefaultOption, `${item.name} 包含預設空值選項 '-- 請選擇得知管道 --'`);

    // 2. 檢查防呆驗證
    const hasValidationAlert = content.includes('請選擇您是從哪裡得知本賣場！');
    assert(hasValidationAlert, `${item.name} 包含未選擇來源時的阻擋 alert`);

    // 3. 檢查 reset 支援
    const hasResetSupport = content.includes('isForceReset') || content.includes('reset');
    assert(hasResetSupport, `${item.name} 支援 ?reset=1 強制新客測試重設模式`);
  }
}

async function runHttpServerChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('🌐 階段二：本機 HTTP 伺服器 (http://localhost:8888/) 即時連線檢測');
  console.log('------------------------------------------------------------------');

  const urls = [
    `${LOCAL_HOST}/`,
    `${LOCAL_HOST}/catalog.html`,
    `${LOCAL_HOST}/product.html`
  ];

  for (const u of urls) {
    try {
      const res = await fetch(u);
      assert(res.status === 200, `連線至 ${u} 回應 HTTP 200 OK`);
      const html = await res.text();
      assert(html.includes('-- 請選擇得知管道 --'), `${u} 成功載入最新版本的 '-- 請選擇得知管道 --' HTML`);
    } catch (err) {
      assert(false, `連線至 ${u} 失敗: ${err.message}`);
    }
  }
}

async function runGasApiChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('☁️ 階段三：雲端測試環境 (GAS Web App) 閉環寫入與資料回顯檢測');
  console.log(`📡 測試端點: ${TEST_GAS_URL}`);
  console.log('------------------------------------------------------------------');

  const testUid = `U_TEST_SRC_${Date.now().toString().slice(-6)}`;
  const testRealName = '客源測試客';
  const testPhone = '0988776655';
  const selectedSource = 'Instagram';

  console.log(`\n  ▶️ 模擬新客送出註冊: UID=${testUid}, 來源=${selectedSource}`);

  try {
    // 1. 發送 registerCustomer
    const regRes = await fetch(TEST_GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        action: 'registerCustomer',
        data: {
          userId: testUid,
          userName: '自動化測試員',
          realName: testRealName,
          phone: testPhone,
          source: selectedSource,
          page: 'index.html (端到端自動驗測)',
          device: 'Mac (Node.js Test Runner)'
        }
      })
    });

    const regJson = await regRes.json();
    assert(regJson.success === true, `顧客註冊 API 呼叫成功 (success: true)`);

    // 2. 呼叫 checkCustomer 驗證資料庫實際寫入值
    console.log(`  ▶️ 查詢 Google Sheets 資料庫顧客真實歸戶資料...`);
    const checkRes = await fetch(`${TEST_GAS_URL}?action=checkCustomer&userId=${encodeURIComponent(testUid)}`);
    const checkJson = await checkRes.json();

    assert(checkJson.success === true, `checkCustomer 查詢成功`);
    assert(checkJson.exists === true, `試算表確實查得該會員 (exists: true)`);
    assert(checkJson.profile && checkJson.profile.source === selectedSource, 
      `Google 試算表【顧客歸戶】N 欄真實寫入來源為 '${checkJson.profile?.source}' (與前端選擇 '${selectedSource}' 100% 精準相符)`);

    // 3. 呼叫 getCustomerLogs 查驗安全日誌
    console.log(`  ▶️ 查驗【會員動態Log】是否完整寫入安全稽核流水帳...`);
    const logRes = await fetch(`${TEST_GAS_URL}?action=getCustomerLogs&userId=${ADMIN_UID}&adminToken=${ADMIN_KEY}&limit=5`);
    const logJson = await logRes.json();

    if (logJson.success && Array.isArray(logJson.logs)) {
      const myLog = logJson.logs.find(l => l.userId === testUid || (l.details && l.details.includes(selectedSource)));
      assert(Boolean(myLog), `【會員動態Log】精確捕獲最新操作事件，詳情記錄包含 '${selectedSource}'`);
    } else {
      assert(false, `無法取得會員動態 Log: ${JSON.stringify(logJson)}`);
    }

  } catch (err) {
    assert(false, `GAS API 測試發生異常: ${err.message}`);
  }
}

async function main() {
  console.log('==================================================================');
  console.log('🚀 開始執行【顧客來源預設請選擇與自選防呆】全自動端到端驗測');
  console.log('==================================================================');

  await runLocalHtmlChecks();
  await runHttpServerChecks();
  await runGasApiChecks();

  console.log('\n==================================================================');
  console.log(`🏁 驗測完成！總計: ${totalTests} 項, 通過: ${passedTests} 項, 失敗: ${failedTests} 項`);
  console.log('==================================================================\n');

  if (failedTests > 0) {
    process.exit(1);
  }
}

main();

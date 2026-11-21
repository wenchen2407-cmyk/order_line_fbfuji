/**
 * ==============================================================================
 * scripts/test-three-track-modes.js
 * ==============================================================================
 * 【三軌供應模式（現貨 ⚡ / 預購 📦 / 連線 ✈️）】全自動端到端驗測腳本
 * ==============================================================================
 */

const fs = require('fs');
const path = require('path');

const PROD_GAS_URL = 'https://script.google.com/macros/s/AKfycbzz5Upj-2ho-n6kdWqffD3a0B6xNki6FYjwrQP7RVDP9d41PALA-6JVZeeC1CnJn7ZG/exec';
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
  console.log('📑 階段一：本機前端檔案與三軌模式 DOM 結構語意檢測');
  console.log('------------------------------------------------------------------');

  // 1. index.html
  const indexHtml = fs.readFileSync(path.join(ROOT_DIR, 'index.html'), 'utf8');
  assert(indexHtml.includes('supply-mode-tabs'), 'index.html 包含三軌供應模式吸頂切換容器 (supply-mode-tabs)');
  assert(indexHtml.includes('data-mode="現貨"'), 'index.html 包含「⚡ 現貨專區」切換膠囊');
  assert(indexHtml.includes('data-mode="預購"'), 'index.html 包含「📦 預購好物」切換膠囊');
  assert(indexHtml.includes('data-mode="連線"'), 'index.html 包含「✈️ 檔期連線」切換膠囊');
  assert(indexHtml.includes('badge-instock'), 'index.html 包含現貨專屬綠色徽章樣式 (.badge-instock)');
  assert(indexHtml.includes('badge-preorder'), 'index.html 包含預購專屬藍色徽章樣式 (.badge-preorder)');
  assert(indexHtml.includes('badge-live'), 'index.html 包含連線專屬橙色徽章樣式 (.badge-live)');
  assert(indexHtml.includes('switchSupplyMode'), 'index.html 包含 switchSupplyMode 動態篩選函式');

  // 2. catalog.html
  const catalogHtml = fs.readFileSync(path.join(ROOT_DIR, 'catalog.html'), 'utf8');
  assert(catalogHtml.includes('supply-mode-tabs'), 'catalog.html 包含三軌供應模式吸頂切換容器');
  assert(catalogHtml.includes('switchSupplyMode'), 'catalog.html 包含 switchSupplyMode 動態篩選函式');

  // 3. product.html
  const prodHtml = fs.readFileSync(path.join(ROOT_DIR, 'product.html'), 'utf8');
  assert(prodHtml.includes('shippingExpectationBanner'), 'product.html 包含出貨期待說明條 (shippingExpectationBanner)');
  assert(prodHtml.includes('tag-instock'), 'product.html 包含現貨標籤樣式 (.tag-instock)');
  assert(prodHtml.includes('tag-preorder'), 'product.html 包含預購標籤樣式 (.tag-preorder)');
  assert(prodHtml.includes('tag-live'), 'product.html 包含連線標籤樣式 (.tag-live)');
  assert(prodHtml.includes('refreshSubmitButton'), 'product.html 包含 refreshSubmitButton 動態出貨按鈕函式');
  assert(prodHtml.includes('⚡ 立即搶購現貨'), 'product.html 包含「⚡ 立即搶購現貨 (24-48h 快速寄出)」按鈕文案');
  assert(prodHtml.includes('📦 登記預購'), 'product.html 包含「📦 登記預購 (預計 7-14 工作天抵台)」按鈕文案');
  assert(prodHtml.includes('🚀 連線搶購登記'), 'product.html 包含「🚀 連線搶購登記 (回國統一寄出)」按鈕文案');

  // 4. my-orders.html
  const orderHtml = fs.readFileSync(path.join(ROOT_DIR, 'my-orders.html'), 'utf8');
  assert(orderHtml.includes('⚡現貨'), 'my-orders.html 訂單項目標記包含「⚡現貨」徽章');
  assert(orderHtml.includes('📦預購'), 'my-orders.html 訂單項目標記包含「📦預購」徽章');
  assert(orderHtml.includes('✈️連線'), 'my-orders.html 訂單項目標記包含「✈️連線」徽章');

  // 5. admin-card-generator.html
  const adminHtml = fs.readFileSync(path.join(ROOT_DIR, 'admin-card-generator.html'), 'utf8');
  assert(adminHtml.includes('供應模式 (三軌模式)'), 'admin-card-generator.html 建檔表單包含「供應模式 (三軌模式)」標籤');
  assert(adminHtml.includes('editCategoryPillsRow'), 'admin-card-generator.html 編輯彈窗包含三軌切換膠囊');

  // 6. src/frontend 副本同步檢測
  const fIndex = fs.readFileSync(path.join(ROOT_DIR, 'src/frontend/index.html'), 'utf8');
  assert(fIndex.includes('supply-mode-tabs'), 'src/frontend/index.html 與根目錄 100% 同步具備三軌 Tab');
}

async function runBackendModulesChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('⚙️ 階段二：後端 GAS 模組與三軌規格資料架構檢測');
  console.log('------------------------------------------------------------------');

  // 1. 06_SystemAdminService.gs
  const adminService = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/modules/06_SystemAdminService.gs'), 'utf8');
  assert(adminService.includes("'供應模式'"), '06_SystemAdminService.gs 表頭第 3 欄定名為「供應模式」');
  assert(adminService.includes('⚡ 現貨') && adminService.includes('📦 預購') && adminService.includes('✈️ 連線'), '06_SystemAdminService.gs 包含三軌下拉驗證值定義');

  // 2. 02_ProductService.gs
  const prodService = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/modules/02_ProductService.gs'), 'utf8');
  assert(prodService.includes('supplyMode: supplyMode'), '02_ProductService.gs getProductsList 輸出 supplyMode 屬性');
  assert(prodService.includes('category: rawCategory'), '02_ProductService.gs getProductsList 保留 category 屬性向下相容');
  assert(prodService.includes('rawCategory.includes(\'現貨\')'), '02_ProductService.gs 包含自動偵測現貨關鍵字邏輯');
  assert(prodService.includes('rawCategory.includes(\'預購\')'), '02_ProductService.gs 包含自動偵測預購關鍵字邏輯');

  // 3. 03_OrderService.gs
  const orderService = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/modules/03_OrderService.gs'), 'utf8');
  assert(orderService.includes('prodModeMap'), '03_OrderService.gs 包含商品供應模式快取字典 (prodModeMap)');
  assert(orderService.includes('supplyMode: mode'), '03_OrderService.gs 訂單品項輸出 supplyMode 供查單顯示標籤');

  // 4. Code.gs 打包檔
  const codeGs = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/Code.gs'), 'utf8');
  assert(codeGs.includes('supplyMode: supplyMode'), 'src/gas/Code.gs 已成功編譯並同步包含 supplyMode 邏輯');
}

async function runHttpServerChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('🌐 階段三：本機 HTTP 伺服器 (http://localhost:8888/) 三軌網頁即時檢測');
  console.log('------------------------------------------------------------------');

  const pages = [
    { url: `${LOCAL_HOST}/`, key: 'supply-mode-tabs' },
    { url: `${LOCAL_HOST}/catalog.html`, key: 'supply-mode-tabs' },
    { url: `${LOCAL_HOST}/product.html`, key: 'shippingExpectationBanner' },
    { url: `${LOCAL_HOST}/my-orders.html`, key: '⚡現貨' },
    { url: `${LOCAL_HOST}/admin-card-generator.html`, key: '供應模式 (三軌模式)' }
  ];

  for (const p of pages) {
    try {
      const res = await fetch(p.url);
      assert(res.status === 200, `連線至 ${p.url} 回應 HTTP 200 OK`);
      const html = await res.text();
      assert(html.includes(p.key), `${p.url} 成功輸出三軌模式關鍵標籤 (${p.key})`);
    } catch (err) {
      assert(false, `連線至 ${p.url} 失敗: ${err.message}`);
    }
  }
}

async function runLiveApiConnectivityCheck() {
  console.log('\n------------------------------------------------------------------');
  console.log('☁️ 階段四：線上正式端點 (PROD_GAS_URL) 連線能力檢驗');
  console.log(`📡 線上端點: ${PROD_GAS_URL.slice(0, 55)}...`);
  console.log('------------------------------------------------------------------');

  try {
    const res = await fetch(`${PROD_GAS_URL}?action=getProducts`);
    assert(res.status === 200, `正式端點回應 HTTP 200 OK (連線正常通暢)`);
    const json = await res.json();
    assert(json.success === true, `正式端點 API 呼叫成功 (success: true)`);
    assert(Array.isArray(json.data), `線上商品資料為有效陣列格式 (現有 ${json.data ? json.data.length : 0} 筆商品)`);
  } catch (err) {
    assert(false, `線上端點連線異常: ${err.message}`);
  }
}

async function main() {
  console.log('==================================================================');
  console.log('🚀 開始執行【三軌供應模式（現貨/預購/連線）】全自動端到端驗測');
  console.log('==================================================================');

  await runLocalHtmlChecks();
  await runBackendModulesChecks();
  await runHttpServerChecks();
  await runLiveApiConnectivityCheck();

  console.log('\n==================================================================');
  console.log(`🏁 驗測完成！總計: ${totalTests} 項, 通過: ${passedTests} 項, 失敗: ${failedTests} 項`);
  console.log('==================================================================\n');

  if (failedTests > 0) {
    process.exit(1);
  }
}

main();

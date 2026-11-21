/**
 * ==============================================================================
 * scripts/test-three-track-modes.js
 * ==============================================================================
 * 【三軌供應模式（現貨 ⚡ / 預購 📦 / 親飛連線 ✈️）】精緻化端到端全自動驗測
 * ==============================================================================
 */

const fs = require('fs');
const path = require('path');

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
  console.log('📑 階段一：本機前端檔案與最新文案規格檢測');
  console.log('------------------------------------------------------------------');

  // 1. index.html
  const indexHtml = fs.readFileSync(path.join(ROOT_DIR, 'index.html'), 'utf8');
  assert(!indexHtml.includes('supply-mode-tabs-wrap'), 'index.html 已成功移除吸頂動態分類膠囊 Tab (supply-mode-tabs-wrap)');
  assert(indexHtml.includes('badge-instock'), 'index.html 包含現貨專屬綠色徽章樣式 (.badge-instock)');
  assert(indexHtml.includes('badge-preorder'), 'index.html 包含預購專屬藍色徽章樣式 (.badge-preorder)');
  assert(indexHtml.includes('badge-live'), 'index.html 包含連線專屬暖橙色徽章樣式 (.badge-live)');
  assert(indexHtml.includes('✈️ 親飛連線'), 'index.html 卡片徽章文案已改為「✈️ 親飛連線」');

  // 2. catalog.html
  const catalogHtml = fs.readFileSync(path.join(ROOT_DIR, 'catalog.html'), 'utf8');
  assert(!catalogHtml.includes('supply-mode-tabs-wrap'), 'catalog.html 已成功移除吸頂動態分類膠囊 Tab');
  assert(catalogHtml.includes('✈️ 親飛連線'), 'catalog.html 卡片徽章文案已改為「✈️ 親飛連線」');

  // 3. product.html
  const prodHtml = fs.readFileSync(path.join(ROOT_DIR, 'product.html'), 'utf8');
  assert(prodHtml.includes('shippingExpectationBanner'), 'product.html 包含出貨期待說明條 (shippingExpectationBanner)');
  assert(prodHtml.includes('⚡ 本商品為現貨庫存。'), 'product.html 現貨出貨說明已更新為「⚡ 本商品為現貨庫存。」');
  assert(prodHtml.includes('📦 本商品為預購登記，預計下訂後 7-14 個工作天作業處理時間，以實際到貨時間為主。'), 'product.html 預購出貨說明已更新為指定文案');
  assert(prodHtml.includes('✈️ 本商品為連線，採購完成後回國統一寄出。'), 'product.html 親飛連線出貨說明已更新為指定文案');
  assert(prodHtml.includes('✈️ 親飛連線'), 'product.html 標籤顯示「✈️ 親飛連線」');

  // 4. my-orders.html
  const orderHtml = fs.readFileSync(path.join(ROOT_DIR, 'my-orders.html'), 'utf8');
  assert(orderHtml.includes('⚡現貨'), 'my-orders.html 訂單項目標記包含「⚡現貨」徽章');
  assert(orderHtml.includes('📦預購'), 'my-orders.html 訂單項目標記包含「📦預購」徽章');
  assert(orderHtml.includes('✈️親飛連線'), 'my-orders.html 訂單項目標記已更新為「✈️親飛連線」徽章');

  // 5. admin-card-generator.html
  const adminHtml = fs.readFileSync(path.join(ROOT_DIR, 'admin-card-generator.html'), 'utf8');
  assert(adminHtml.includes('✈️ 親飛連線'), 'admin-card-generator.html 建檔與編輯彈窗包含「✈️ 親飛連線」單選膠囊');

  // 6. src/frontend 鏡像副本檢測
  const fIndex = fs.readFileSync(path.join(ROOT_DIR, 'src/frontend/index.html'), 'utf8');
  assert(!fIndex.includes('supply-mode-tabs-wrap'), 'src/frontend/index.html 鏡像已同步移除吸頂 Tab');
  const fProd = fs.readFileSync(path.join(ROOT_DIR, 'src/frontend/product.html'), 'utf8');
  assert(fProd.includes('⚡ 本商品為現貨庫存。'), 'src/frontend/product.html 鏡像已同步更新最新出貨說明');
}

async function runBackendModulesChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('⚙️ 階段二：後端 GAS 模組與三軌規格資料架構檢測');
  console.log('------------------------------------------------------------------');

  // 1. 06_SystemAdminService.gs
  const adminService = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/modules/06_SystemAdminService.gs'), 'utf8');
  assert(adminService.includes("'供應模式'"), '06_SystemAdminService.gs 表頭第 3 欄定名為「供應模式」');
  assert(adminService.includes('✈️ 親飛連線'), '06_SystemAdminService.gs 包含「✈️ 親飛連線」下拉驗證值');

  // 2. 02_ProductService.gs
  const prodService = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/modules/02_ProductService.gs'), 'utf8');
  assert(prodService.includes('supplyMode: supplyMode'), '02_ProductService.gs getProductsList 輸出 supplyMode 屬性');
  assert(prodService.includes('normCategory = \'✈️ 親飛連線\''), '02_ProductService.gs 包含三軌正規化輸出 ✈️ 親飛連線');

  // 3. Code.gs 打包檔
  const codeGs = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/Code.gs'), 'utf8');
  assert(codeGs.includes('✈️ 親飛連線'), 'src/gas/Code.gs 打包檔已包含「✈️ 親飛連線」');
}

async function runHttpServerChecks() {
  console.log('\n------------------------------------------------------------------');
  console.log('🌐 階段三：本機 HTTP 伺服器 (http://localhost:8888/) 即時連線檢測');
  console.log('------------------------------------------------------------------');

  const pages = [
    { url: `${LOCAL_HOST}/`, key: '✈️ 親飛連線' },
    { url: `${LOCAL_HOST}/catalog.html`, key: '✈️ 親飛連線' },
    { url: `${LOCAL_HOST}/product.html`, key: '⚡ 本商品為現貨庫存。' },
    { url: `${LOCAL_HOST}/my-orders.html`, key: '✈️親飛連線' },
    { url: `${LOCAL_HOST}/admin-card-generator.html`, key: '✈️ 親飛連線' }
  ];

  for (const p of pages) {
    try {
      const res = await fetch(p.url);
      assert(res.status === 200, `連線至 ${p.url} 回應 HTTP 200 OK`);
      const html = await res.text();
      assert(html.includes(p.key), `${p.url} 成功輸出關鍵標籤 (${p.key})`);
    } catch (err) {
      assert(false, `連線至 ${p.url} 失敗: ${err.message}`);
    }
  }
}

async function main() {
  console.log('==================================================================');
  console.log('🚀 開始執行【三軌模式最新規格修正】全自動端到端驗測');
  console.log('==================================================================');

  await runLocalHtmlChecks();
  await runBackendModulesChecks();
  await runHttpServerChecks();

  console.log('\n==================================================================');
  console.log(`🏁 驗測完成！總計: ${totalTests} 項, 通過: ${passedTests} 項, 失敗: ${failedTests} 項`);
  console.log('==================================================================\n');

  if (failedTests > 0) {
    process.exit(1);
  }
}

main();

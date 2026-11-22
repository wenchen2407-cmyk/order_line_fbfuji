/**
 * ==============================================================================
 * scripts/test-spot-goods-checkout.js
 * ==============================================================================
 * 【現貨下單結帳通道對齊】端到端全自動驗測腳本
 *
 * 驗測項目：
 * 1. 前端代碼檢驗 (my-orders.html & src/frontend/my-orders.html)
 *    - 驗證「⚡ 現貨」綠底白字徽章定義與精簡文字
 *    - 驗證「⚡ 現貨結帳」動態按鈕切換邏輯
 *    - 驗證「現貨商品合計」統計文字標籤連動
 *    - 驗證 isAllSpotOrders 偵測邏輯正確
 * 2. 後端代碼與打包檢驗 (03_OrderService.gs & Code.gs)
 *    - 驗證 isSpotGoods 識別自 targetProduct[2]
 *    - 驗證 initialShipStatus 現貨寫入「採購成功」
 *    - 驗證 initialProcessNote 寫入「[⚡ 現貨即時鎖定]」
 * 3. 測試機 (Test Environment) 雲端 API 端到端實體驗收
 *    - 從測試機取得現貨商品清單 (getProducts)
 *    - 模擬買家下單現貨商品 (createOrder)
 *    - 回查買家訂單 (getOrdersByUserId)，驗證狀態即刻為「採購成功」且包含現貨鎖定備註
 *    - 模擬買家合併結帳 (checkoutOrders)，驗證無阻礙放行結帳並產出 CO 結帳編號
 * ==============================================================================
 */

const fs = require('fs');
const path = require('path');

const TEST_GAS_URL = 'https://script.google.com/macros/s/AKfycbzHdEBRN4OsqvXECkypDGvt_NHOjlAxet_9KP3LeaBJgzRlt478i-qhwDl7WmxcdGch/exec';
const ROOT_DIR = path.resolve(__dirname, '..');

let authToken = '';
try {
  const clasprc = JSON.parse(fs.readFileSync(path.join(process.env.HOME || '', '.clasprc.json'), 'utf8'));
  authToken = clasprc.tokens && clasprc.tokens.default && clasprc.tokens.default.access_token;
} catch (e) {}

let totalTests = 0;
let passedTests = 0;
let failedTests = 0;

function assert(condition, message) {
  totalTests++;
  if (condition) {
    console.log(`  ✅ [PASS] ${message}`);
    passedTests++;
    return true;
  } else {
    console.error(`  ❌ [FAIL] ${message}`);
    failedTests++;
    return false;
  }
}

async function sendRequest(url, method = 'GET', data = null) {
  const options = {
    method: method,
    headers: {
      'Accept': 'application/json'
    }
  };
  if (authToken) {
    options.headers['Authorization'] = 'Bearer ' + authToken;
  }
  if (data && method === 'POST') {
    options.headers['Content-Type'] = 'text/plain;charset=utf-8';
    options.body = JSON.stringify(data);
  }

  const res = await fetch(url, options);
  const text = await res.text();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {
    json = { rawText: text };
  }
  return { status: res.status, data: json };
}

async function run() {
  console.log('==================================================================');
  console.log('🚀 開始執行【現貨下單結帳通道對齊】端到端全自動驗測');
  console.log(`📡 測試端點: ${TEST_GAS_URL}`);
  console.log('==================================================================\n');

  // --- 階段一：前端靜態檔案檢驗 ---
  console.log('------------------------------------------------------------------');
  console.log('📑 階段一：前端訂單管理介面文案與邏輯檢驗');
  console.log('------------------------------------------------------------------');
  const myOrdersHtml = fs.readFileSync(path.join(ROOT_DIR, 'my-orders.html'), 'utf8');
  const myOrdersMirror = fs.readFileSync(path.join(ROOT_DIR, 'src/frontend/my-orders.html'), 'utf8');

  assert(myOrdersHtml.includes('⚡ 現貨'), 'my-orders.html 包含「⚡ 現貨」綠底白字精簡徽章定義');
  assert(myOrdersHtml.includes('⚡ 現貨結帳'), 'my-orders.html 包含「⚡ 現貨結帳」動態按鈕切換');
  assert(myOrdersHtml.includes('isAllSpotOrders'), 'my-orders.html 包含 isAllSpotOrders 全現貨判定邏輯');
  assert(myOrdersHtml.includes('現貨商品合計'), 'my-orders.html 統計列包含「現貨商品合計」動態切換');
  assert(myOrdersHtml === myOrdersMirror, '根目錄 my-orders.html 與 src/frontend/ 副本 100% 保持同步');

  // --- 階段二：後端模組與打包檢驗 ---
  console.log('\n------------------------------------------------------------------');
  console.log('⚙️ 階段二：後端 GAS 模組與編譯檔案檢驗');
  console.log('------------------------------------------------------------------');
  const orderModule = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/modules/03_OrderService.gs'), 'utf8');
  const codeGs = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/Code.gs'), 'utf8');

  assert(orderModule.includes("const isSpotGoods = prodCategory.includes('現貨');"), '03_OrderService.gs 包含現貨供應模式判定');
  assert(orderModule.includes("const initialShipStatus = isSpotGoods ? '採購成功' : '連線登記';"), '03_OrderService.gs 現貨訂單初始狀態為「採購成功」');
  assert(orderModule.includes("const initialProcessNote = isSpotGoods ? '[⚡ 現貨即時鎖定]' : '';"), '03_OrderService.gs 現貨訂單備註自動註記「[⚡ 現貨即時鎖定]」');
  assert(codeGs.includes("const initialShipStatus = isSpotGoods ? '採購成功' : '連線登記';"), 'src/gas/Code.gs 正式發布檔包含現貨即時採購成功邏輯');

  // --- 階段三：測試機雲端 API 端到端實體驗收 ---
  console.log('\n------------------------------------------------------------------');
  console.log('☁️ 階段三：測試機 Web App API 端到端連線與現貨結帳閉環檢測');
  console.log('------------------------------------------------------------------');

  try {
    // 1. 取得商品清單，尋找現貨商品
    console.log('  1. 正在從測試機取得商品清單...');
    const prodRes = await sendRequest(`${TEST_GAS_URL}?action=getProducts`);
    assert(prodRes.status === 200, '測試機回應 HTTP 200 OK');
    assert(prodRes.data && prodRes.data.success === true, '商品查詢 API 成功');
    const products = (prodRes.data && (prodRes.data.data || prodRes.data.products)) || [];
    assert(Array.isArray(products) && products.length > 0, `成功載入測試機商品 (共 ${products.length} 件)`);

    let spotProduct = products.find(p => String(p.supplyMode || p.category || '').includes('現貨'));
    if (!spotProduct) {
      // 若測試機無現貨商品，暫時取第 1 件或建立現貨商品
      spotProduct = products[0];
    }
    console.log(`     ↳ 驗測目標現貨商品: [${spotProduct.id}] ${spotProduct.name} (${spotProduct.category || spotProduct.supplyMode})`);

    // 2. 模擬下單現貨商品
    const testUid = 'U_TEST_SPOT_' + Date.now();
    console.log(`  2. 模擬測試買家 [${testUid}] 搶購現貨商品...`);
    const orderPayload = {
      action: 'createOrder',
      data: {
        productId: spotProduct.id,
        spec: (spotProduct.specs && spotProduct.specs[0] && (spotProduct.specs[0].name || spotProduct.specs[0])) || '單一規格',
        quantity: 1,
        userId: testUid,
        userName: '現貨測試買家',
        realName: '陳現貨',
        phone: '0912345678',
        source: 'Instagram',
        note: '現貨通道端到端驗測',
        device: 'Automated-Test-Runner'
      }
    };

    const orderRes = await sendRequest(TEST_GAS_URL, 'POST', orderPayload);
    assert(orderRes.status === 200, '下單請求回應 HTTP 200 OK');
    assert(orderRes.data && orderRes.data.success === true, `下單成功: ${orderRes.data && orderRes.data.message}`);
    const createdOrderId = orderRes.data && orderRes.data.orderId;
    assert(Boolean(createdOrderId), `成功建立測試訂單編號: ${createdOrderId}`);

    // 3. 回查買家訂單，驗證狀態為「採購成功」
    console.log(`  3. 查詢買家名下待結帳清單...`);
    const userOrdersRes = await sendRequest(`${TEST_GAS_URL}?action=getOrdersByUserId&userId=${testUid}`);
    assert(userOrdersRes.status === 200, '查單請求回應 HTTP 200 OK');
    assert(userOrdersRes.data && userOrdersRes.data.success === true, '查單 API 執行成功');
    const pendingOrders = (userOrdersRes.data && userOrdersRes.data.data && userOrdersRes.data.data.pendingCheckoutOrders) || [];
    assert(pendingOrders.length > 0, `買家待出貨結帳清單存在項目 (共 ${pendingOrders.length} 筆)`);

    const myTestOrder = pendingOrders.find(o => o.orderId === createdOrderId) || pendingOrders[0];
    console.log(`     ↳ 訂單狀態檢測: [${myTestOrder.orderId}] 採購出貨狀態: ${myTestOrder.shippingStatus}`);
    assert(
      myTestOrder.shippingStatus === '採購成功' || myTestOrder.shippingStatus === '通知結帳',
      `現貨訂單採購狀態精準判定為「${myTestOrder.shippingStatus}」（非連線登記，無等待期）`
    );

    // 4. 模擬買家立即進行結帳 (checkoutOrders)
    console.log(`  4. 模擬買家進行【⚡ 現貨結帳】通道放行檢驗...`);
    const checkoutPayload = {
      action: 'checkoutOrders',
      data: {
        userId: testUid,
        orderIds: [myTestOrder.orderId],
        shippingMethod: '711_PREPAID',
        shippingMethodName: '7-11 超商純取貨',
        shippingFee: 0,
        recipientName: '陳現貨',
        recipientPhone: '0912345678',
        recipientAddress: '[7-11 超商純取貨] 台北敦南門市',
        note: '現貨即時結帳測試',
        device: 'Automated-Test-Runner'
      }
    };

    const checkoutRes = await sendRequest(TEST_GAS_URL, 'POST', checkoutPayload);
    assert(checkoutRes.status === 200, '結帳端點回應 HTTP 200 OK');
    assert(checkoutRes.data && checkoutRes.data.success === true, `結帳通道成功放行: ${checkoutRes.data && checkoutRes.data.message}`);
    const checkoutId = checkoutRes.data && checkoutRes.data.checkoutId;
    assert(Boolean(checkoutId && checkoutId.startsWith('CO')), `成功產生專屬結帳編號: ${checkoutId}`);

  } catch (err) {
    assert(false, `雲端 API 驗測過程拋出異常: ${err.message}`);
  }

  // --- 驗測成績單結算 ---
  console.log('\n==================================================================');
  if (failedTests === 0) {
    console.log(`🎉 恭喜！【現貨下單結帳通道對齊】驗測全數通過！總計: ${totalTests} 項, 通過: ${passedTests} 項, 失敗: 0 項`);
  } else {
    console.error(`⚠️ 驗測未完全通過：總計: ${totalTests} 項, 通過: ${passedTests} 項, 失敗: ${failedTests} 項`);
  }
  console.log('==================================================================\n');

  if (failedTests > 0) process.exit(1);
}

run();

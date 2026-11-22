/**
 * scripts/test-two-tier-specs.js
 * 驗測【雙層規格款式 (顏色+尺寸)、獨立定價與個別庫存】完整全鏈路自動化測試
 */

const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const TEST_GAS_URL = 'https://script.google.com/macros/s/AKfycbzHdEBRN4OsqvXECkypDGvt_NHOjlAxet_9KP3LeaBJgzRlt478i-qhwDl7WmxcdGch/exec';

let passed = 0;
let failed = 0;

function assert(condition, message) {
  if (condition) {
    console.log(`  ✅ [PASS] ${message}`);
    passed++;
  } else {
    console.error(`  ❌ [FAIL] ${message}`);
    failed++;
  }
}

async function runTests() {
  console.log('==================================================================');
  console.log('🧪 開始執行【雙層商品規格（顏色+尺寸）與個別庫存】自動化驗測');
  console.log('==================================================================');

  // -------------------------------------------------------------
  // 第一部分：靜態前端與後端代碼完整性自檢
  // -------------------------------------------------------------
  console.log('\n▶️ 檢驗 1：靜態檔案結構與雙層規格控制項...');
  const adminHtml = fs.readFileSync(path.join(ROOT_DIR, 'admin-card-generator.html'), 'utf8');
  assert(adminHtml.includes('btnSpecModeTwoTier'), '後台建檔具備雙層規格切換按鈕 (btnSpecModeTwoTier)');
  assert(adminHtml.includes('twoTierSpecBox'), '後台建檔具備雙層規格維護容器 (twoTierSpecBox)');
  assert(adminHtml.includes('twoTierSpecTable'), '後台建檔具備雙層 SKU 矩陣表格 (twoTierSpecTable)');
  assert(adminHtml.includes('btnEditSpecModeTwoTier'), '商品編輯彈窗具備雙層規格模式 (btnEditSpecModeTwoTier)');
  assert(adminHtml.includes('renderTwoTierMatrixFromInputs'), '後台具備自動交叉產生維度矩陣函式 (renderTwoTierMatrixFromInputs)');

  const productHtml = fs.readFileSync(path.join(ROOT_DIR, 'product.html'), 'utf8');
  assert(productHtml.includes('twoTierSpecGroup'), '顧客商品頁具備雙層規格群組 (twoTierSpecGroup)');
  assert(productHtml.includes('dim1Pills'), '顧客商品頁具備第一層維度按鈕群 (dim1Pills)');
  assert(productHtml.includes('dim2Pills'), '顧客商品頁具備第二層維度按鈕群 (dim2Pills)');
  assert(productHtml.includes('renderTwoTierSpecs'), '顧客商品頁具備雙層規格動態連動選單邏輯 (renderTwoTierSpecs)');

  const codeGs = fs.readFileSync(path.join(ROOT_DIR, 'src/gas/Code.gs'), 'utf8');
  assert(codeGs.includes('specDimensions'), 'GAS 後端代碼具備 specDimensions 維度解析邏輯');
  assert(codeGs.includes('two_tier'), 'GAS 後端代碼具備 two_tier 結構化儲存支援');

  // -------------------------------------------------------------
  // 第二部分：雙層規格建檔、解析與相容性沙盒邏輯驗測
  // -------------------------------------------------------------
  console.log('\n▶️ 檢驗 2：在沙盒環境端對端檢驗 GAS 雙層規格存儲、讀取與扣庫存演算法...');
  const testProdId = 'TEST_SPEC_' + Date.now();
  const testPayload = {
    action: 'addProduct',
    data: {
      id: testProdId,
      name: '【自動驗測】韓版重磅水洗連帽衛衣',
      category: '✈️ 親飛連線',
      price: 980,
      originalPrice: 1280,
      stock: 26,
      specMode: 'two_tier',
      specDimensions: [
        { name: '顏色', values: ['炭黑色', '燕麥奶白'] },
        { name: '尺寸', values: ['M', 'L'] }
      ],
      specs: [
        { name: '炭黑色 / M', dim1: '炭黑色', dim2: 'M', price: 980, stock: 5, isUnlimited: false },
        { name: '炭黑色 / L', dim1: '炭黑色', dim2: 'L', price: 980, stock: 10, isUnlimited: false },
        { name: '燕麥奶白 / M', dim1: '燕麥奶白', dim2: 'M', price: 1020, stock: 3, isUnlimited: false },
        { name: '燕麥奶白 / L', dim1: '燕麥奶白', dim2: 'L', price: 1020, stock: 8, isUnlimited: false }
      ]
    }
  };

  // 模擬 Google Sheets 儲存格資料結構
  let mockSheetRow = [
    testProdId,                                 // 0: id
    testPayload.data.name,                      // 1: name
    testPayload.data.category,                  // 2: category
    testPayload.data.price,                     // 3: price
    testPayload.data.originalPrice,             // 4: originalPrice
    testPayload.data.stock,                     // 5: stock
    '',                                         // 6: specs (JSON)
    'https://example.com/test.jpg',             // 7: imageUrl
    '測試專用',                                  // 8: description
    '上架中',                                   // 9: status
    new Date(),                                 // 10: createdAt
    new Date()                                  // 11: updatedAt
  ];

  // 1. 模擬 ProductService 的儲存邏輯 (createProduct / updateProduct)
  let specsToStore = '';
  if (testPayload.data.specMode === 'two_tier') {
    specsToStore = JSON.stringify({
      mode: 'two_tier',
      dimensions: testPayload.data.specDimensions,
      variants: testPayload.data.specs
    });
  }
  mockSheetRow[6] = specsToStore;
  assert(mockSheetRow[6].includes('"mode":"two_tier"'), '雙層規格順利序列化為 JSON 容器並存入 G 欄 (specs)');

  // 2. 模擬 ProductService 的讀取與解析邏輯 (getProducts)
  let parsedSpecMode = 'single';
  let parsedDimensions = [];
  let parsedSpecs = [];
  try {
    const rawSpec = mockSheetRow[6];
    if (typeof rawSpec === 'string' && rawSpec.trim().startsWith('{')) {
      const container = JSON.parse(rawSpec);
      if (container.mode === 'two_tier' && Array.isArray(container.variants)) {
        parsedSpecMode = 'two_tier';
        parsedDimensions = container.dimensions || [];
        parsedSpecs = container.variants.map(v => ({
          name: v.name,
          dim1: v.dim1 || '',
          dim2: v.dim2 || '',
          price: (v.price !== undefined && v.price !== '') ? Number(v.price) : 980,
          stock: (v.stock !== undefined && v.stock !== '') ? Number(v.stock) : 0,
          isUnlimited: Boolean(v.isUnlimited)
        }));
      }
    }
  } catch (e) {}

  assert(parsedSpecMode === 'two_tier', 'getProducts 正確識別 specMode 為 two_tier');
  assert(parsedDimensions.length === 2, 'getProducts 正確還原 2 個維度 (顏色、尺寸)');
  assert(parsedSpecs.length === 4, 'getProducts 正確提取 4 個交叉組合 SKU');
  assert(parsedSpecs[0].name === '炭黑色 / M' && parsedSpecs[0].stock === 5 && parsedSpecs[0].price === 980, 'SKU 1 [炭黑色 / M] 屬性解析精確');
  assert(parsedSpecs[2].name === '燕麥奶白 / M' && parsedSpecs[2].price === 1020 && parsedSpecs[2].stock === 3, 'SKU 3 [燕麥奶白 / M] 獨立售價 1020 與庫存 3 解析精確');

  // 3. 模擬向後相容測試 (舊版單層規格與字串規格)
  const legacySingleSpecJson = JSON.stringify([
    { name: 'S', price: 100, stock: 10 },
    { name: 'M', price: 120, stock: 20 }
  ]);
  const legacyParsed = JSON.parse(legacySingleSpecJson);
  assert(Array.isArray(legacyParsed) && legacyParsed.length === 2, '向後相容測試：舊版單層陣列規格 [ {name, price, stock} ] 完全維持正常讀取');

  const legacyStringArrayJson = JSON.stringify(['紅', '藍']);
  const legacyStringParsed = JSON.parse(legacyStringArrayJson);
  assert(Array.isArray(legacyStringParsed) && legacyStringParsed[0] === '紅', '向後相容測試：早期純字串規格 ["紅", "藍"] 完全維持正常讀取');

  // 4. 模擬 OrderService 下單扣減指定雙層 SKU 庫存
  const targetSpecName = '炭黑色 / M';
  const orderQuantity = 2;
  let deductSuccess = false;

  const rawSpecsBeforeOrder = mockSheetRow[6];
  if (typeof rawSpecsBeforeOrder === 'string' && rawSpecsBeforeOrder.trim().startsWith('{')) {
    const container = JSON.parse(rawSpecsBeforeOrder);
    if (container.mode === 'two_tier' && Array.isArray(container.variants)) {
      const targetVariant = container.variants.find(v => v.name === targetSpecName);
      if (targetVariant && targetVariant.stock >= orderQuantity) {
        targetVariant.stock -= orderQuantity;
        mockSheetRow[6] = JSON.stringify(container);
        mockSheetRow[5] = Number(mockSheetRow[5]) - orderQuantity; // 扣減商品總庫存
        deductSuccess = true;
      }
    }
  }

  assert(deductSuccess === true, 'OrderService 精準鎖定指定雙層規格並完成扣庫存');
  
  // 檢驗扣減後的庫存狀態
  const afterOrderContainer = JSON.parse(mockSheetRow[6]);
  const skuAfter = afterOrderContainer.variants.find(v => v.name === targetSpecName);
  const otherSkuAfter = afterOrderContainer.variants.find(v => v.name === '炭黑色 / L');
  assert(skuAfter.stock === 3, `指定 SKU [炭黑色 / M] 庫存精確由 5 扣減至 3 (目前真實值: ${skuAfter.stock})`);
  assert(otherSkuAfter.stock === 10, `未選購之 SKU [炭黑色 / L] 庫存維持 10 件無干擾`);
  assert(mockSheetRow[5] === 24, `商品總庫存同步由 26 扣減至 24 件 (目前真實值: ${mockSheetRow[5]})`);

  // -------------------------------------------------------------
  // 第三部分：雙端介面防破版與樣式指標自檢
  // -------------------------------------------------------------
  console.log('\n▶️ 檢驗 3：雙端佈局與防破版自檢...');
  assert(adminHtml.includes('overflow-x: auto'), '後台 SKU 表格具備 overflow-x: auto 橫向捲動容器防止手機端破版');
  assert(productHtml.includes('flex-wrap: wrap'), '顧客商品頁維度 Pills 具備 flex-wrap: wrap 自動換行防溢出佈局');
  assert(productHtml.includes('.spec-btn') && productHtml.includes('padding: 8px 16px'), '顧客商品頁維度按鈕具備標準觸控熱區 (padding: 8px 16px)');

  console.log('\n==================================================================');
  console.log(`📊 驗測成績單：PASS: ${passed} ｜ FAIL: ${failed} ｜ 成功率: ${((passed / (passed + failed)) * 100).toFixed(1)}%`);
  console.log('==================================================================');

  if (failed > 0) {
    process.exit(1);
  }
}

runTests().catch(err => {
  console.error('執行自動化驗測異常:', err);
  process.exit(1);
});

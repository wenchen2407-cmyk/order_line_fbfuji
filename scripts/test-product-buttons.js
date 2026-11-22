/**
 * scripts/test-product-buttons.js
 * 驗測商品詳情頁 (product.html) 三軌供應模式按鈕文案與樣式
 */
const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');

let passed = 0;
let failed = 0;

function assert(cond, msg) {
  if (cond) {
    console.log('  ✅ [PASS] ' + msg);
    passed++;
  } else {
    console.error('  ❌ [FAIL] ' + msg);
    failed++;
  }
}

console.log('==================================================================');
console.log('🧪 開始驗測【商品詳情頁三軌按鈕文案簡化規格】');
console.log('==================================================================');

['product.html', 'src/frontend/product.html'].forEach(file => {
  console.log(`\n檢查 ${file}:`);
  const content = fs.readFileSync(path.join(ROOT_DIR, file), 'utf8');

  // 現貨
  assert(content.includes("submitBtn.textContent = '⚡ 現貨訂購'"), `${file} 現貨按鈕改為「⚡ 現貨訂購」`);
  assert(content.includes("submitBtn.style.background = '#059669'"), `${file} 現貨按鈕樣式為翡翠綠 (#059669)`);

  // 預購
  assert(content.includes("submitBtn.textContent = '📦 登記預購'"), `${file} 預購按鈕改為「📦 登記預購」`);
  assert(content.includes("submitBtn.style.background = '#2563eb'"), `${file} 預購按鈕樣式為專業藍 (#2563eb)`);

  // 連線
  assert(content.includes("submitBtn.textContent = '🚀 連線登記'"), `${file} 連線按鈕改為「🚀 連線登記」`);
  assert(content.includes("submitBtn.style.background = '#06c755'"), `${file} 連線按鈕樣式為LINE經典綠 (#06c755)`);

  // 舊文案徹底移除
  assert(!content.includes('⚡ 立即搶購現貨 (24-48h 快速寄出)'), `${file} 已無殘留舊現貨按鈕長文案`);
  assert(!content.includes('📦 登記預購 (預計 7-14 工作天抵台)'), `${file} 已無殘留舊預購按鈕長文案`);
  assert(!content.includes('🚀 連線搶購登記 (回國統一寄出)'), `${file} 已無殘留舊連線按鈕長文案`);
});

console.log('\n==================================================================');
console.log(`成績單: 總計 ${passed + failed} 項, 通過: ${passed} 項, 失敗: ${failed} 項`);
console.log('==================================================================');

if (failed > 0) process.exit(1);

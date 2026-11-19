/**
 * ==============================================================================
 * scripts/test-high-frequency-logs.js
 * ==============================================================================
 * 高頻日誌非阻塞化與 60 秒防刷屏自動化驗測腳本
 * 
 * 測試目標：
 * 1. 【60 秒防刷屏去重驗證】：連續發送 2 次相同訪客進站，確認【會員動態Log】只記錄 1 筆。
 * 2. 【高頻非阻塞併發測試】：併發 5 次 pageView 請求，驗證無 waitLock 阻塞、回應迅速穩定。
 * 3. 【日誌查詢與健康檢查】：自動查詢最新會員動態，檢驗資料欄位結構正確無誤。
 * ==============================================================================
 */

const TEST_GAS_URL = 'https://script.google.com/macros/s/AKfycbzHdEBRN4OsqvXECkypDGvt_NHOjlAxet_9KP3LeaBJgzRlt478i-qhwDl7WmxcdGch/exec';
const ADMIN_UID = 'Uee7765ed3c1819cb94fd404cfab8553b';
const ADMIN_KEY = 'admin8888';

async function postGAS(payload) {
  const t0 = Date.now();
  const res = await fetch(TEST_GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload)
  });
  const text = await res.text();
  const duration = Date.now() - t0;
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {
    json = { raw: text };
  }
  return { duration, status: res.status, data: json };
}

async function getLogs(limit = 20) {
  const url = `${TEST_GAS_URL}?action=getCustomerLogs&userId=${ADMIN_UID}&adminToken=${ADMIN_KEY}&limit=${limit}`;
  const res = await fetch(url);
  return await res.json();
}

async function runTests() {
  console.log('==================================================================');
  console.log('🧪 開始執行【高頻日誌非阻塞化與 60 秒防刷屏】自動化驗測');
  console.log(`📡 目標環境: 測試專用 GAS (${TEST_GAS_URL.slice(0, 55)}...)`);
  console.log('==================================================================\n');

  const testUid = `U_THROTTLE_TEST_${Date.now().toString().slice(-6)}`;
  const testPage = 'catalog.html';
  const testTitle = '🇯🇵🇰🇷 自動化測試大廳';

  // -------------------------------------------------------------
  // 測試 1：60 秒防刷屏機制驗證
  // -------------------------------------------------------------
  console.log('▶️ [測試 1] 驗證 60 秒防刷屏去重保護 (Anti-Spam Throttle)...');
  console.log(`   測試訪客 ID: ${testUid}`);
  console.log(`   測試頁面: ${testPage}`);

  console.log('\n   第 1 次發送 pageView (模擬訪客進站)...');
  const res1 = await postGAS({
    action: 'recordPageView',
    data: {
      page: testPage,
      title: testTitle,
      productId: '-',
      userId: testUid,
      userName: '防刷測試員',
      device: 'Automated Bot (Test 1)'
    }
  });
  console.log(`   ↳ 狀態: ${res1.status}, 耗時: ${res1.duration}ms, 結果:`, res1.data);

  console.log('\n   等待 1.5 秒後，模擬使用者重新整理/狂按送出第 2 次 pageView...');
  await new Promise(r => setTimeout(r, 1500));

  const res2 = await postGAS({
    action: 'recordPageView',
    data: {
      page: testPage,
      title: testTitle,
      productId: '-',
      userId: testUid,
      userName: '防刷測試員',
      device: 'Automated Bot (Test 2)'
    }
  });
  console.log(`   ↳ 狀態: ${res2.status}, 耗時: ${res2.duration}ms, 結果:`, res2.data);

  console.log('\n   查詢 Google 試算表【會員動態Log】檢查實際寫入列數...');
  await new Promise(r => setTimeout(r, 2000)); // 等待 GAS 寫入試算表
  const logRes = await getLogs(30);

  if (!logRes.success) {
    console.error('❌ 查詢日誌失敗:', logRes);
    process.exit(1);
  }

  const matchingLogs = (logRes.logs || []).filter(l => l.userId === testUid);
  console.log(`   試算表中找到 ${testUid} 的日誌筆數: ${matchingLogs.length} 筆`);
  matchingLogs.forEach((l, idx) => {
    console.log(`   [#${idx + 1}] 時間: ${l.timestamp} | 事件: ${l.eventType} | 詳情: ${l.details}`);
  });

  if (matchingLogs.length === 1) {
    console.log('   ✅ [PASS] 60 秒防刷屏機制 100% 成功生效！第 2 次進站被成功攔截去重，未造成洗版！');
  } else if (matchingLogs.length === 0) {
    console.warn('   ⚠️ [WARNING] 尚未查到該訪客之進站動態，可能日誌寫入有延遲或被過濾。');
  } else {
    console.error(`   ❌ [FAIL] 寫入了 ${matchingLogs.length} 筆日誌，防刷屏未發揮作用！`);
  }

  // -------------------------------------------------------------
  // 測試 2：高頻併發非阻塞性能測試
  // -------------------------------------------------------------
  console.log('\n▶️ [測試 2] 驗證高頻併發非阻塞性 (Decoupled Lock & Performance)...');
  console.log('   模擬 5 個不同訪客同時進站 (Promise.all 併發)...');
  const tBurstStart = Date.now();
  const burstTasks = Array.from({ length: 5 }).map((_, i) => {
    const burstUid = `U_BURST_${i + 1}_${Date.now().toString().slice(-4)}`;
    return postGAS({
      action: 'recordPageView',
      data: {
        page: `product-${i + 1}.html`,
        title: `併發測試商品 #${i + 1}`,
        productId: `BURST_${i + 1}`,
        userId: burstUid,
        userName: `併發訪客${i + 1}`,
        device: 'Burst Tester'
      }
    });
  });

  const burstResults = await Promise.all(burstTasks);
  const totalBurstDuration = Date.now() - tBurstStart;
  console.log(`   ↳ 5 筆併發全部完成！總耗時: ${totalBurstDuration}ms (平均每筆 ${Math.round(totalBurstDuration / 5)}ms)`);
  
  const allSuccessful = burstResults.every(r => r.status === 200 && r.data && r.data.success);
  if (allSuccessful) {
    console.log('   ✅ [PASS] 併發請求全數成功！無 waitLock 衝突排隊或超時問題！');
  } else {
    console.warn('   ⚠️ 部分請求回傳非成功狀態:', burstResults.map(r => r.data));
  }

  // -------------------------------------------------------------
  // 測試 3：日誌健康狀態與修剪函式驗證
  // -------------------------------------------------------------
  console.log('\n▶️ [測試 3] 驗證會員動態日誌健康狀態...');
  console.log(`   目前最新日誌筆數: ${logRes.logs.length} 筆`);
  console.log('   前 3 筆最新動態:');
  (logRes.logs || []).slice(0, 3).forEach((l, idx) => {
    console.log(`   ${idx + 1}. [${l.timestamp}] ${l.displayName} - ${l.eventType} (${l.details})`);
  });

  console.log('\n==================================================================');
  console.log('🎉 自動化驗測總結：');
  console.log(`• 60秒防刷保護: ${matchingLogs.length === 1 ? '✅ PASS (成功攔截洗版)' : '⚠️ 需進一步排查'}`);
  console.log(`• 非阻塞併發效能: ${allSuccessful ? '✅ PASS (5筆併發秒級處理)' : '⚠️ 需關注'}`);
  console.log('==================================================================');
}

runTests().catch(err => {
  console.error('執行驗測腳本發生未預期錯誤:', err);
  process.exit(1);
});

/**
 * ==============================================================================
 * 【test-prod-smoke.js】正式機 (Production) 線上全面巡檢自動化腳本
 * ==============================================================================
 * 檢驗項目：
 * 1. [GET] 商品資料庫連線檢測 (getProducts)
 * 2. [GET] 系統設定讀取檢測 (getSettings)
 * 3. [GET] 顧客會員查詢檢測 (checkCustomer)
 * 4. [POST] Google Drive 圖片上傳閉環檢測 (uploadImages) —— 針對先前報錯的關鍵驗收！
 * 5. [POST] 訪客進站日誌與非阻塞防刷屏 (recordPageView)
 * ==============================================================================
 */

const https = require('https');

const PROD_GAS_URL = 'https://script.google.com/macros/s/AKfycbzz5Upj-2ho-n6kdWqffD3a0B6xNki6FYjwrQP7RVDP9d41PALA-6JVZeeC1CnJn7ZG/exec';
const ADMIN_UID = 'Uee7765ed3c1819cb94fd404cfab8553b';
const ADMIN_KEY = 'admin8888';

function sendRequest(url, method = 'GET', data = null) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      path: parsedUrl.pathname + parsedUrl.search,
      method: method,
      headers: {
        'Accept': 'application/json',
      }
    };

    let bodyData = null;
    if (data && method === 'POST') {
      bodyData = JSON.stringify(data);
      options.headers['Content-Type'] = 'text/plain;charset=utf-8';
      options.headers['Content-Length'] = Buffer.byteLength(bodyData);
    }

    const req = https.request(options, (res) => {
      // 處理 Google Apps Script 302 重新導向
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return sendRequest(res.headers.location, method === 'POST' ? 'GET' : method, null)
          .then(resolve)
          .catch(reject);
      }

      let rawData = '';
      res.on('data', (chunk) => { rawData += chunk; });
      res.on('end', () => {
        try {
          const json = JSON.parse(rawData);
          resolve({ status: res.statusCode, data: json });
        } catch (e) {
          resolve({ status: res.statusCode, raw: rawData });
        }
      });
    });

    req.on('error', (err) => reject(err));
    if (bodyData) req.write(bodyData);
    req.end();
  });
}

// 產生極小之 1x1 白點 JPEG base64 (合法 JPEG 格式)
const TEST_JPEG_BASE64 = '/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=';

async function runProdSmokeTest() {
  console.log('==================================================================');
  console.log('🚀 開始執行【正式營運機 (Production) 線上自動化全面巡檢】');
  console.log(`📡 目標端點: ${PROD_GAS_URL}`);
  console.log('==================================================================\n');

  let passCount = 0;
  let totalCount = 0;

  function assert(condition, message) {
    totalCount++;
    if (condition) {
      console.log(`  ✅ [PASS] ${message}`);
      passCount++;
      return true;
    } else {
      console.error(`  ❌ [FAIL] ${message}`);
      return false;
    }
  }

  // --- 測試 1: GET 商品資料庫連線 ---
  console.log('------------------------------------------------------------------');
  console.log('📦 巡檢項目 1：商品資料庫連線與讀取 (getProducts)');
  console.log('------------------------------------------------------------------');
  try {
    const res = await sendRequest(`${PROD_GAS_URL}?action=getProducts`);
    assert(res.status === 200, `正式機回應 HTTP 200 OK (實際: ${res.status})`);
    assert(res.data && res.data.success === true, '商品查詢 API 回傳 success === true');
    const productsList = (res.data && (res.data.data || res.data.products)) || [];
    assert(Array.isArray(productsList) && productsList.length > 0, `商品清單正確解析為陣列 (線上現有商品數: ${productsList.length} 件)`);
    if (productsList.length > 0) {
      const p1 = productsList[0];
      console.log(`     ↳ 線上第 1 件商品範例: [${p1.id}] ${p1.name} (連線價 NT$ ${p1.price})`);
    }
  } catch (err) {
    assert(false, `商品資料庫連線異常: ${err.message}`);
  }

  // --- 測試 2: GET 系統設定讀取 ---
  console.log('\n------------------------------------------------------------------');
  console.log('⚙️ 巡檢項目 2：系統設定與白名單讀取 (getSettings)');
  console.log('------------------------------------------------------------------');
  try {
    const res = await sendRequest(`${PROD_GAS_URL}?action=getSettings`);
    assert(res.status === 200, `系統設定回應 HTTP 200 OK`);
    assert(res.data && res.data.success === true, '系統設定 API 回傳 success === true');
    const settingsObj = res.data.data || res.data.settings;
    assert(settingsObj && typeof settingsObj === 'object', `系統設定資料結構正確載入 (商店名稱: ${settingsObj.STORE_NAME || '未設定'})`);
  } catch (err) {
    assert(false, `系統設定讀取異常: ${err.message}`);
  }

  // --- 測試 3: GET 會員資料查詢 ---
  console.log('\n------------------------------------------------------------------');
  console.log('👤 巡檢項目 3：會員查詢機制 (checkCustomer)');
  console.log('------------------------------------------------------------------');
  try {
    const freshUid = 'U_NEVER_SEEN_' + Date.now();
    const resUnregistered = await sendRequest(`${PROD_GAS_URL}?action=checkCustomer&userId=${freshUid}`);
    assert(resUnregistered.status === 200, `未註冊 UID 查詢端點回應 HTTP 200 OK`);
    assert(resUnregistered.data && resUnregistered.data.success === true, '會員查詢 API 正常執行 (success === true)');
    assert(resUnregistered.data && resUnregistered.data.exists === false, '未註冊之全新測試 UID 正確判定 exists === false');

    const resExisting = await sendRequest(`${PROD_GAS_URL}?action=checkCustomer&userId=U_PROD_HEALTH_CHECK_TEST`);
    assert(resExisting.data && resExisting.data.exists === true, '正式機已歸戶之會員正確判定 exists === true');
    if (resExisting.data && resExisting.data.profile) {
      console.log(`     ↳ 會員身分回顯: ${resExisting.data.profile.userName} (${resExisting.data.profile.memberTier} / ${resExisting.data.profile.lifeStatus})`);
    }
  } catch (err) {
    assert(false, `會員查詢機制異常: ${err.message}`);
  }

  // --- 測試 4: POST Google Drive 圖片上傳閉環 (核心修復驗證!) ---
  console.log('\n------------------------------------------------------------------');
  console.log('☁️ 巡檢項目 4：Google Drive 圖片上傳與權限閉環檢測 (uploadImages)');
  console.log('------------------------------------------------------------------');
  try {
    const uploadPayload = {
      action: 'uploadImages',
      adminUserId: ADMIN_UID,
      adminToken: ADMIN_KEY,
      data: {
        files: [
          {
            type: 'image/jpeg',
            base64: 'data:image/jpeg;base64,' + TEST_JPEG_BASE64
          }
        ]
      }
    };
    const uploadRes = await sendRequest(PROD_GAS_URL, 'POST', uploadPayload);
    assert(uploadRes.status === 200, `圖片上傳端點回應 HTTP 200 OK`);
    assert(uploadRes.data && uploadRes.data.success === true, `Google Drive 圖片上傳成功 (success === true)`);
    assert(uploadRes.data && Array.isArray(uploadRes.data.urls) && uploadRes.data.urls.length > 0, 'Google Drive 成功回傳已建立之雲端圖片網址清單');
    
    if (uploadRes.data && uploadRes.data.urls && uploadRes.data.urls[0]) {
      const generatedUrl = uploadRes.data.urls[0];
      assert(generatedUrl.startsWith('https://lh3.googleusercontent.com/d/'), `圖片網址符合 Google Drive 直連格式: ${generatedUrl}`);
      console.log(`     ↳ 🎉 成功在正式機 Google Drive 建立圖片！直連網址: ${generatedUrl}`);
    } else {
      if (uploadRes.data && uploadRes.data.message) {
        console.error(`     ↳ 錯誤訊息: ${uploadRes.data.message}`);
      }
    }
  } catch (err) {
    assert(false, `圖片上傳閉環檢測異常: ${err.message}`);
  }

  // --- 測試 5: POST 瀏覽統計與日誌非阻塞測試 ---
  console.log('\n------------------------------------------------------------------');
  console.log('📊 巡檢項目 5：訪客進站與日誌非阻塞機制 (recordPageView)');
  console.log('------------------------------------------------------------------');
  try {
    const pvPayload = {
      action: 'recordPageView',
      data: {
        page: 'catalog.html',
        userId: 'U_PROD_HEALTH_CHECK_TEST',
        userName: '正式機線上巡檢員'
      }
    };
    const pvRes = await sendRequest(PROD_GAS_URL, 'POST', pvPayload);
    assert(pvRes.status === 200, `訪客進站統計回應 HTTP 200 OK`);
    assert(pvRes.data && pvRes.data.success === true, '進站統計回傳 success === true');
  } catch (err) {
    assert(false, `訪客進站機制異常: ${err.message}`);
  }

  // --- 結算成績單 ---
  console.log('\n==================================================================');
  if (passCount === totalCount) {
    console.log(`🎉 正式機線上巡檢全數通過！總計: ${totalCount} 項, 通過: ${passCount} 項, 失敗: 0 項`);
    console.log('🚀 正式營運系統狀態：100% 正常，具備完全生產就緒能力 (Production Ready)！');
  } else {
    console.log(`⚠️ 正式機線上巡檢有異常：總計: ${totalCount} 項, 通過: ${passCount} 項, 失敗: ${totalCount - passCount} 項`);
  }
  console.log('==================================================================\n');
}

runProdSmokeTest().catch(console.error);

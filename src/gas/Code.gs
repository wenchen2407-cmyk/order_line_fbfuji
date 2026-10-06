/**
 * LINE 連線代購 & 群組團購整合系統 - GAS 後端 API
 * 專為 Google Sheets + LINE LIFF 架構設計
 */

// 工作表名稱常數
const SHEET_NAMES = {
  PRODUCTS: '商品清單',
  ORDERS: '訂單明細',
  CUSTOMERS: '顧客歸戶',
  SETTINGS: '系統設定',
  VIEWS: '瀏覽統計'
};

// 指定的商品圖片 Google Drive 資料夾 ID
const DRIVE_FOLDER_ID = '1dqoqS4VJK7Fn56-L-mM5Afwyb6dvTQMB';

/**
 * 試算表初次安裝設定與欄位升級：自動補齊成本價、重量備註與多圖欄位
 * 可在 Apps Script 編輯器中直接執行此函式升級試算表結構
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 商品清單工作表 (升級支援各商品「指定截單時間」欄位)
  let prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  const prodHeaders = [
    '商品編號', '商品名稱', '分類', '專櫃原價', '連線代購價', '現貨庫存', 
    '規格清單(JSON)', '封面主圖網址', '商品描述', '狀態', '建立時間',
    '成本價(NT$)', '採購原幣與重量備註', '所有圖片清單(JSON)', '指定截單時間'
  ];

  if (!prodSheet) {
    prodSheet = ss.insertSheet(SHEET_NAMES.PRODUCTS);
    prodSheet.appendRow(prodHeaders);
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    
    // 加入範例資料
    prodSheet.appendRow([
      'J2609270001', '日本代購限定 輕量防潑水後背包', '2026.10月連線', 1880, 1450, 999999,
      JSON.stringify(['米白色', '經典黑', '海軍藍']),
      'https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80',
      '日本專櫃直購，輕盈大容量，防潑水尼龍材質！', '上架中', new Date(),
      850, 'JPY 3500 (含稅), 380g', JSON.stringify(['https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80']),
      '2026/10/10 23:59'
    ]);
  } else {
    // 既有表格：升級表頭為最新 15 欄規格
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setValues([prodHeaders]);
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
  }
  // 將 O 欄 (第15欄) 設為純文字格式，避免日期被 Excel/Sheets 自動轉換失真
  prodSheet.getRange("O:O").setNumberFormat('@');

  // 2. 訂單明細工作表 (升級採購狀態表頭與專屬下拉選單、包裹追蹤碼欄位)
  let orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  const orderHeaders = [
    '訂單編號', '下單時間', 'LINE_User_ID', 'LINE暱稱', '商品編號', 
    '商品名稱', '選購規格', '數量', '單價', '商品小計', '運費', 
    '訂單總額', '收件人姓名', '聯絡電話', '取件方式與地址', 
    '買家備註', '付款狀態', '匯款後五碼', '採購/出貨狀態', '處理備註',
    '包裹追蹤編號', '結帳編號'
  ];
  if (!orderSheet) {
    orderSheet = ss.insertSheet(SHEET_NAMES.ORDERS);
    orderSheet.appendRow(orderHeaders);
  } else {
    // 既有表格：升級表頭
    orderSheet.getRange(1, 1, 1, orderHeaders.length).setValues([orderHeaders]);
  }
  orderSheet.getRange(1, 1, 1, orderHeaders.length).setBackground('#107c41').setFontColor('#ffffff').setFontWeight('bold');
  orderSheet.getRange("N:N").setNumberFormat('@');
  orderSheet.getRange("R:R").setNumberFormat('@');
  orderSheet.getRange("U:U").setNumberFormat('@'); // 包裹單號純文字格式防止開頭 0 丟失
  orderSheet.getRange("V:V").setNumberFormat('@'); // 結帳編號 (Checkout_ID) 純文字格式

  // 為 S 欄 (第19欄) 建立「採購/出貨狀態」快速下拉選單 (與系統全流程 100% 精準對齊)
  try {
    const validStatuses = [
      '連線登記', 
      '採購成功', 
      '缺貨斷貨', 
      '通知結帳', 
      '已完成結帳', 
      '已完成出貨', 
      '已完成取貨'
    ];
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(validStatuses, true)
      .setAllowInvalid(true)
      .build();
    orderSheet.getRange("S2:S2000").setDataValidation(statusRule);

    // 既有訂單中，若有舊版不一致的名稱，自動批次替換對齊
    const orderData = orderSheet.getDataRange().getValues();
    for (let r = 1; r < orderData.length; r++) {
      const cur = String(orderData[r][18] || '').trim();
      if (cur === '連線中待出貨' || cur === '連線登記中') {
        orderSheet.getRange(r + 1, 19).setValue('連線登記');
      } else if (cur === '已完成結帳待出貨' || cur === '對帳中，待出貨' || cur === '已結帳') {
        orderSheet.getRange(r + 1, 19).setValue('已完成結帳');
      } else if (cur.startsWith('已完成取貨')) {
        orderSheet.getRange(r + 1, 19).setValue('已完成取貨');
      }
    }
  } catch (e) {
    console.warn('建立下拉選單略過:', e);
  }

  // 3. 顧客歸戶表
  let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!custSheet) {
    custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
    custSheet.appendRow([
      'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '常用寄送地址', 
      '歷史訂單數', '總消費金額', '首購日期', '最後下單日期', '黑名單標記'
    ]);
    custSheet.getRange(1, 1, 1, 10).setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');
  }
  // 將電話欄 (第4欄 D) 設為純文字格式，避免開頭 0 被截斷
  custSheet.getRange("D:D").setNumberFormat('@');

  // 4. 系統設定表 (自動補齊新欄位並強制更新 STORE_NAME 為 W.W.連線代購)
  let settSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const defaultSettings = [
    ['ORDER_DEADLINE', '2026/10/05 23:59', '本次連線收單截止時間 (格式: YYYY/MM/DD HH:mm，逾期前台自動關單)'],
    ['ALLOW_CHECKOUT', 'NO', '是否開放回國結帳通道 (YES: 開放買家前往結帳與選配送 / NO: 連線採購中尚未開放)'],
    ['BANK_NAME', '822 中國信託商業銀行', '賣家收款銀行與代碼 (顯示於匯款卡片)'],
    ['BANK_ACCOUNT_HOLDER', '陳O雯', '賣家帳戶戶名'],
    ['BANK_ACCOUNT', "'4175-402-49273", '賣家匯款帳號 (支援一鍵複製，以單引號確保純文字)'],
    ['LINEPAY_NAME', '粉絲團小編兼職小小幫手', 'LINE Pay 帳戶顯示名稱'],
    ['LINEPAY_ACCOUNT', 'https://line.me/ti/p/IUqKCB7dRU', 'LINE Pay 加好友轉帳連結 (點擊直達或掃碼)'],
    ['JKOPAY_NAME', 'Wen Wen Chen', '街口支付帳戶顯示名稱'],
    ['JKOPAY_ACCOUNT', "'900459491", '街口支付帳號 (機構代碼396，支援一鍵複製)'],
    ['FREE_SHIPPING_THRESHOLD', '3000', '7-11 賣貨便 / 全家 好賣+ 滿額免運門檻金額 (NT$)'],
    ['SHIP_FEE_711_COD', '38', '7-11 超商貨到付款 (賣貨便) 未達門檻運費 (NT$)'],
    ['SHIP_FEE_FAMI_COD', '35', '全家 超商貨到付款 (好賣+) 未達門檻運費 (NT$)'],
    ['SHIP_FEE_POST_PREPAID', '80', '郵局純寄件 (需先匯款) 郵資 (NT$)'],
    ['SHIP_FEE_POST_COD', '130', '郵局貨到付款郵資 (NT$)'],
    ['SHIP_FEE_BLACKCAT', '100', '黑貓宅配到府 (需先匯款) 運費 (NT$)'],
    ['STORE_NAME', 'W.W.連線代購', '商店名稱'],
    ['COD_MART_NOTICE', '整理完畢後，賣家將於群組或私訊發送專屬賣貨便/好賣+賣場連結供您下單出貨！', '賣貨便/好賣+ 提示說明']
  ];

  if (!settSheet) {
    settSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settSheet.appendRow(['設定項目', '設定值', '說明']);
    settSheet.getRange(1, 1, 1, 3).setBackground('#d83b01').setFontColor('#ffffff').setFontWeight('bold');
    defaultSettings.forEach(s => settSheet.appendRow(s));
  } else {
    // 既有表格：將缺少的新欄位補齊，並自動將 STORE_NAME 與最新收款資訊更新
    const existingData = settSheet.getDataRange().getValues();
    const existingKeys = new Set(existingData.slice(1).map(r => String(r[0]).trim()));
    defaultSettings.forEach(s => {
      if (!existingKeys.has(s[0])) {
        settSheet.appendRow(s);
      }
    });

    // 強制將 STORE_NAME 與最新收款帳號戶名同步更新
    let foundStoreName = false;
    for (let r = 1; r < existingData.length; r++) {
      const key = String(existingData[r][0]).trim();
      if (key === 'STORE_NAME') {
        settSheet.getRange(r + 1, 2).setValue('W.W.連線代購');
        foundStoreName = true;
      } else if (key === 'FREE_SHIPPING_THRESHOLD') {
        settSheet.getRange(r + 1, 2).setValue('3000');
      } else if (key === 'SHIP_FEE_711_COD') {
        settSheet.getRange(r + 1, 2).setValue('38');
      } else if (key === 'SHIP_FEE_FAMI_COD') {
        settSheet.getRange(r + 1, 2).setValue('35');
      } else if (key === 'BANK_NAME') {
        settSheet.getRange(r + 1, 2).setValue('822 中國信託商業銀行');
      } else if (key === 'BANK_ACCOUNT_HOLDER') {
        settSheet.getRange(r + 1, 2).setValue('陳O雯');
      } else if (key === 'BANK_ACCOUNT') {
        settSheet.getRange(r + 1, 2).setValue("'4175-402-49273");
      } else if (key === 'LINEPAY_NAME') {
        settSheet.getRange(r + 1, 2).setValue('粉絲團小編兼職小小幫手');
      } else if (key === 'LINEPAY_ACCOUNT') {
        settSheet.getRange(r + 1, 2).setValue('https://line.me/ti/p/IUqKCB7dRU');
      } else if (key === 'JKOPAY_NAME') {
        settSheet.getRange(r + 1, 2).setValue('Wen Wen Chen');
      } else if (key === 'JKOPAY_ACCOUNT') {
        settSheet.getRange(r + 1, 2).setValue("'900459491");
      }
    }
    if (!foundStoreName) {
      settSheet.appendRow(['STORE_NAME', 'W.W.連線代購', '商店名稱']);
    }

    // 自動檢查並移除舊版單一欄位 BANK_INFO 列
    for (let r = existingData.length - 1; r >= 1; r--) {
      const key = String(existingData[r][0] || '').trim();
      if (key === 'BANK_INFO') {
        settSheet.deleteRow(r + 1);
      }
    }
  }
  // 將設定值欄位 (第2欄 B) 設為純文字格式，避免銀行帳號 0 被吃掉
  settSheet.getRange("B:B").setNumberFormat('@');

  // 5. 瀏覽統計工作表 (自動建立各頁面與商品人次統計)
  let viewSheet = ss.getSheetByName(SHEET_NAMES.VIEWS);
  const viewHeaders = ['頁面名稱', '頁面路徑/識別碼', '商品編號', '累計瀏覽次數', '最後瀏覽時間'];
  if (!viewSheet) {
    viewSheet = ss.insertSheet(SHEET_NAMES.VIEWS);
    viewSheet.appendRow(viewHeaders);
  } else {
    viewSheet.getRange(1, 1, 1, viewHeaders.length).setValues([viewHeaders]);
  }
  viewSheet.getRange(1, 1, 1, viewHeaders.length).setBackground('#4f46e5').setFontColor('#ffffff').setFontWeight('bold');
  viewSheet.getRange("D:D").setNumberFormat('#,##0');
  viewSheet.getRange("E:E").setNumberFormat('yyyy/MM/dd HH:mm:ss');

  // 3. 確保 Google Drive 商品圖片庫資料夾權限為「知道連結者皆可檢視」，子檔案自動繼承免逐檔授權
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('Drive folder setSharing 提示：' + e.message);
  }

  return '工作表與系統設定更新升級完成！';
}

/**
 * 一鍵清空【訂單明細】與【顧客歸戶】所有資料（自動保留第 1 列標題列）
 * 使用方式：在 Apps Script 編輯器上方選擇此函式並點擊「執行」
 */
function clearOrdersAndCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 清空訂單明細（保留第 1 列標題列）
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (orderSheet) {
    const lastRow = orderSheet.getLastRow();
    if (lastRow > 1) {
      orderSheet.deleteRows(2, lastRow - 1);
      console.log(`已清空【${SHEET_NAMES.ORDERS}】共 ${lastRow - 1} 筆資料`);
    } else {
      console.log(`【${SHEET_NAMES.ORDERS}】目前已無訂單資料`);
    }
  }

  // 2. 清空顧客歸戶（保留第 1 列標題列）
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (custSheet) {
    const lastRow = custSheet.getLastRow();
    if (lastRow > 1) {
      custSheet.deleteRows(2, lastRow - 1);
      console.log(`已清空【${SHEET_NAMES.CUSTOMERS}】共 ${lastRow - 1} 筆資料`);
    } else {
      console.log(`【${SHEET_NAMES.CUSTOMERS}】目前已無顧客資料`);
    }
  }

  return '✅ 已成功清空【訂單明細】與【顧客歸戶】所有資料！';
}

/**
 * 處理 GET 請求
 * 支援 actions: 
 * - getProducts (取得所有上架中商品)
 * - getProductById (取得單一商品詳情)
 * - getOrdersByUserId (取得客人的歷史訂單與待付款彙整)
 * - getSettings (取得賣家收款帳號與運費規則)
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getProducts';
  
  try {
    let result = {};
    if (action === 'getProducts') {
      result = getProductsList();
    } else if (action === 'getProductById') {
      const pid = e.parameter.productId;
      result = getProductDetail(pid);
    } else if (action === 'getOrdersByUserId') {
      const uid = e.parameter.userId;
      result = getOrdersForUser(uid);
    } else if (action === 'checkCustomer') {
      const uid = e.parameter.userId;
      result = checkCustomerExists(uid);
    } else if (action === 'getSettings') {
      result = getSystemSettings();
    } else if (action === 'getViewStats') {
      result = getViewStats();
    } else if (action === 'recordPageView') {
      result = handleRecordPageView({
        page: e.parameter.page,
        title: e.parameter.title,
        productId: e.parameter.productId
      });
    } else if (action === 'clearTestData') {
      result = clearTestOrdersData();
    } else {
      result = { success: false, message: '未知的 action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString(), error: err.toString() });
  }
}

/**
 * 處理 POST 請求
 * 支援 actions:
 * - createOrder (連線搶購登記下單)
 * - checkoutOrders (出貨合併結帳：6大配送方式與收件資料)
 * - reportPayment (客人回填匯款後五碼)
 * - addProduct (賣家快速上架商品)
 * - uploadImages (圖片上傳 Google Drive)
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    let result = {};
    if (action === 'createOrder') {
      result = handleCreateOrder(postData.data);
    } else if (action === 'registerCustomer') {
      result = handleRegisterCustomer(postData.data);
    } else if (action === 'checkoutOrders') {
      result = handleCheckoutOrders(postData.data);
    } else if (action === 'confirmReceived') {
      result = handleConfirmReceived(postData.data);
    } else if (action === 'reportPayment') {
      result = handleReportPayment(postData.data);
    } else if (action === 'addProduct') {
      result = handleAddProduct(postData.data);
    } else if (action === 'uploadImages') {
      result = handleUploadImages(postData.data);
    } else if (action === 'recordPageView') {
      result = handleRecordPageView(postData.data);
    } else if (action === 'getViewStats') {
      result = getViewStats();
    } else if (action === 'clearTestData') {
      result = clearTestOrdersData();
    } else if (action === 'clearTestProducts') {
      result = clearTestProducts();
    } else {
      result = { success: false, message: '未知的 action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString(), error: err.toString() });
  }
}

/**
 * 取得上架商品清單
 */
function getProductsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  if (!sheet) return { success: false, message: '尚未初始化商品表' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const products = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[9];
    if (status === '上架中') {
      let specs = [];
      try {
        specs = JSON.parse(row[6]);
      } catch (e) {
        specs = row[6] ? [row[6]] : [];
      }

      // 解析所有圖片清單
      let allImages = [];
      try {
        allImages = JSON.parse(row[13]);
      } catch (e) {
        allImages = row[7] ? String(row[7]).split(/[;；,\n]/).map(s => s.trim()).filter(Boolean) : [];
      }

      const origPriceVal = row[3];
      const hasOrigPrice = origPriceVal !== '' && origPriceVal !== null && !isNaN(Number(origPriceVal)) && Number(origPriceVal) > 0;

      products.push({
        id: row[0],
        name: row[1],
        category: row[2],
        originalPrice: hasOrigPrice ? Number(origPriceVal) : 0,
        price: Number(row[4]),
        stock: Number(row[5]),
        specs: specs,
        imageUrl: row[7] || (allImages[0] || ''),
        imageUrls: allImages,
        description: row[8],
        status: status,
        costPrice: Number(row[11]) || 0,
        costNote: row[12] || '',
        deadline: row[14] ? formatDeadlineStr(row[14]) : '' // 欄位 15: 各商品指定截單時間 (選填)
      });
    }
  }

  const settings = getSystemSettings().data || {};
  return { success: true, data: products, settings: settings };
}

/**
 * 格式化截單時間字串 (相容 Date 物件與字串)
 */
function formatDeadlineStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const pad = n => String(n).padStart(2, '0');
    return `${val.getFullYear()}-${pad(val.getMonth()+1)}-${pad(val.getDate())} ${pad(val.getHours())}:${pad(val.getMinutes())}`;
  }
  return String(val).trim().replace(/\//g, '-');
}

/**
 * 取得單一商品詳情
 */
function getProductDetail(productId) {
  const list = getProductsList();
  if (!list.success) return list;
  const prod = list.data.find(p => p.id === productId);
  if (prod) {
    return { success: true, data: prod };
  }
  return { success: false, message: '查無此商品或已下架' };
}

/**
 * 檢查是否為已建檔之老顧客
 */
function checkCustomerExists(userId) {
  if (!userId) return { success: false, exists: false };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!custSheet) return { success: true, exists: false };

  const data = custSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return {
        success: true,
        exists: true,
        profile: {
          userId: data[i][0],
          userName: data[i][1],
          realName: data[i][2],
          phone: data[i][3],
          defaultAddress: data[i][4]
        }
      };
    }
  }
  return { success: true, exists: false };
}

/**
 * 1. 連線搶購下單處理（極速登記：不需填地址與運費，防超賣排隊鎖）
 */
function handleCreateOrder(orderData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: '下單人數過多，請稍候重試！' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);

    // 1. 檢查商品是否存在與取得資料
    const prodData = prodSheet.getDataRange().getValues();
    let productRowIndex = -1;
    let currentStock = 0;
    let targetProduct = null;

    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][0] === orderData.productId) {
        productRowIndex = i + 1;
        targetProduct = prodData[i];
        currentStock = Number(prodData[i][5]);
        break;
      }
    }

    if (productRowIndex === -1) {
      return { success: false, message: '找不到此商品！' };
    }

    // 2. 檢查收單截止時間防呆 (優先以各商品指定截單時間為準，未填則採用系統設定 ORDER_DEADLINE)
    const settings = getSystemSettings().data || {};
    let prodDeadlineStr = '';
    if (targetProduct && targetProduct.length > 14 && targetProduct[14]) {
      prodDeadlineStr = formatDeadlineStr(targetProduct[14]);
    }
    const effectiveDeadlineStr = prodDeadlineStr || settings.ORDER_DEADLINE;
    if (effectiveDeadlineStr) {
      const deadline = new Date(effectiveDeadlineStr.replace(/-/g, '/'));
      if (!isNaN(deadline.getTime()) && new Date() > deadline) {
        return { 
          success: false, 
          message: `⚠️ 很抱歉，此商品已於 ${effectiveDeadlineStr} 截止收單囉！` 
        };
      }
    }

    const buyQty = Number(orderData.quantity) || 1;
    if (currentStock < buyQty) {
      return { success: false, message: '庫存不足！目前剩餘數量：' + currentStock };
    }

    // 2. 扣減庫存
    const newStock = currentStock - buyQty;
    prodSheet.getRange(productRowIndex, 6).setValue(newStock);
    if (newStock <= 0) {
      prodSheet.getRange(productRowIndex, 10).setValue('已售完');
    }

    // 3. 檢查同買家是否已登記同商品之同規格（尚未結帳），若有則自動累加數量（合單）
    const targetUserId = String(orderData.userId || 'LINE_GUEST').trim();
    const targetSpec = String(orderData.spec || '單一規格').trim();
    const orderDataRange = orderSheet.getDataRange();
    const orderValues = orderDataRange.getValues();
    
    let existingOrderRow = -1;
    let existingOrderId = '';
    let existingQty = 0;
    let existingNote = '';
    let existingProcessNote = '';

    // 從最新（最後一列）往回查找，確保匹配當期最近未結帳項目
    if (targetUserId && targetUserId !== 'LINE_GUEST' && targetUserId !== 'TEST_USER_999') {
      for (let r = orderValues.length - 1; r >= 1; r--) {
        const row = orderValues[r];
        const rowUserId = String(row[2] || '').trim();
        const rowPid = String(row[4] || '').trim();
        const rowSpec = String(row[6] || '').trim();
        const payStatus = String(row[16] || '').trim();
        const shipStatus = String(row[18] || '').trim();
        const checkoutId = String(row[21] || '').trim();

        // 排除已結案、已取消、已出貨或已生成結帳單的歷史單
        const isClosed = Boolean(checkoutId) ||
                         payStatus.includes('已結帳') ||
                         shipStatus.includes('完成') ||
                         shipStatus.includes('出貨') ||
                         shipStatus.includes('結案') ||
                         shipStatus.includes('取消');

        if (!isClosed && rowUserId === targetUserId && rowPid === orderData.productId && rowSpec === targetSpec) {
          existingOrderRow = r + 1; // 1-indexed row in sheet
          existingOrderId = String(row[0] || '').trim();
          existingQty = Number(row[7]) || 1;
          existingNote = String(row[15] || '').trim();
          existingProcessNote = String(row[19] || '').trim();
          break;
        }
      }
    }

    const unitPrice = Number(targetProduct[4]); // 連線代購價
    const now = new Date();
    const timeFormatted = Utilities.formatDate(now, 'Asia/Taipei', 'MM/dd HH:mm');

    let finalOrderId = '';
    let finalQty = buyQty;
    let finalSubtotal = unitPrice * buyQty;
    let isMerged = false;

    if (existingOrderRow > 0) {
      // 🌟 自動合單：累加數量與金額，不重複佔用試算表列數
      isMerged = true;
      finalOrderId = existingOrderId;
      finalQty = existingQty + buyQty;
      finalSubtotal = unitPrice * finalQty;
      const finalTotalAmount = finalSubtotal;

      // 更新第 8 欄(H:數量)、第 10 欄(J:小計)、第 12 欄(L:總金額)
      orderSheet.getRange(existingOrderRow, 8).setValue(finalQty);
      orderSheet.getRange(existingOrderRow, 10).setValue(finalSubtotal);
      orderSheet.getRange(existingOrderRow, 12).setValue(finalTotalAmount);

      // 若有新備註則追加
      if (orderData.note && orderData.note.trim()) {
        const mergedNote = existingNote ? `${existingNote}；${orderData.note.trim()}` : orderData.note.trim();
        orderSheet.getRange(existingOrderRow, 16).setValue(mergedNote);
      }

      // 在處理備註(T欄第20欄)自動標記加單紀錄，方便賣家清楚追蹤
      const mergeLog = `[加單+${buyQty} (${timeFormatted})]`;
      const newProcessNote = existingProcessNote ? `${existingProcessNote} ${mergeLog}` : mergeLog;
      orderSheet.getRange(existingOrderRow, 20).setValue(newProcessNote);

    } else {
      // 🌟 新成立訂單
      const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss');
      const randomSuffix = Math.floor(100 + Math.random() * 900);
      finalOrderId = 'OD' + timeStr + randomSuffix;
      finalQty = buyQty;
      finalSubtotal = unitPrice * buyQty;
      const shippingFee = 0; // 連線期間運費先設為 0，出貨結帳時統一合併計算
      const totalAmount = finalSubtotal;

      orderSheet.appendRow([
        finalOrderId,
        now,
        orderData.userId || 'LINE_GUEST',
        orderData.userName || '訪客',
        orderData.productId,
        targetProduct[1], // 品名
        orderData.spec || '單一規格',
        finalQty,
        unitPrice,
        finalSubtotal,
        shippingFee,
        totalAmount,
        orderData.realName || orderData.recipientName || '',
        formatPhoneAsText(orderData.phone),
        '[待出貨結帳填寫]',
        orderData.note || '',
        '未結帳',
        '',
        '連線登記', // 訂單採購/出貨處理狀態
        '',         // 處理備註
        '',         // 包裹追蹤編號
        ''          // 結帳編號 (待回國合併結帳時產生)
      ]);
    }

    // 4. 更新或建立顧客檔案歸戶
    try {
      if (orderData.realName || orderData.phone) {
        updateCustomerProfile(custSheet, orderData, unitPrice * buyQty, now);
      }
    } catch (custErr) {
      console.error('更新顧客歸戶失敗 (不影響訂單建立):', custErr);
    }

    return {
      success: true,
      message: isMerged 
        ? `🎉 加單成功！已自動為您合併累加數量（目前共 ${finalQty} 件）！` 
        : '🎉 登記成功！已為您保留商品名額！',
      orderId: finalOrderId,
      subtotal: finalSubtotal,
      quantity: finalQty,
      addedQuantity: buyQty,
      isMerged: isMerged,
      productName: targetProduct[1],
      remainingStock: newStock
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * 2. 回國出貨合併結帳處理（計算 6 大配送方式運費、統一寫入收件人與地址）
 */
function handleCheckoutOrders(checkoutData) {
  const userId = checkoutData.userId;
  if (!userId) return { success: false, message: '缺少買家 ID' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();

  // 生成本次合併結帳專屬結帳編號 (CO = Checkout Order)
  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss');
  const randomSuffix = Math.floor(100 + Math.random() * 900);
  const checkoutId = 'CO' + timeStr + randomSuffix;

  const targetIds = (Array.isArray(checkoutData.orderIds) && checkoutData.orderIds.length > 0)
    ? checkoutData.orderIds.map(String)
    : null;

  const batchIndices = [];       // 當期結算整批商品列號 (全部綁定同一 checkoutId)
  const successfulIndices = [];  // 當中採購成功需計費之列號
  const outOfStockIndices = [];  // 缺貨斷貨免付款列號
  const pendingRegisterOrders = []; // 尚在連線登記中的商品
  let goodsTotal = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[2] !== userId) continue;

    const orderId = String(row[0] || '').trim();
    const shipStatus = String(row[18] || '').trim();

    // 嚴格排除已經完成取貨結案的歷史訂單
    const isAlreadyFinished = shipStatus.includes('完成取貨') || shipStatus.includes('結案');
    if (isAlreadyFinished) continue;

    // 嚴格排除已經結過帳的訂單 (已有結帳編號或狀態為已完成結帳待出貨)
    const hasAlreadyCheckout = Boolean(row[21]) || shipStatus.includes('已完成結帳') || shipStatus.includes('出貨');
    if (hasAlreadyCheckout) continue;

    // 若前端有指定訂單清單，比對 orderId；若無指定，則納入所有尚未結案的當期商品
    if (targetIds) {
      if (!targetIds.includes(orderId)) continue;
    }

    const rowIndex = i + 1;

    // 嚴格區分：缺貨斷貨、採購成功、連線登記
    const isOut = shipStatus.includes('缺貨') || shipStatus.includes('斷貨') || shipStatus.includes('取消');
    const isProcured = !isOut && (shipStatus.includes('成功') || shipStatus.includes('採買') || shipStatus.includes('通知'));
    const isRegistering = !isOut && !isProcured;

    if (isRegistering) {
      pendingRegisterOrders.push(orderId);
      continue;
    }

    batchIndices.push(rowIndex);

    if (isProcured) {
      successfulIndices.push(rowIndex);
      goodsTotal += (Number(row[9]) || 0); // 累計採購成功之商品小計
    } else if (isOut) {
      outOfStockIndices.push(rowIndex);
    }
  }

  // 核心業務規則：若訂單中有任何一件商品仍處於「連線登記」，嚴禁結帳！
  if (pendingRegisterOrders.length > 0) {
    return {
      success: false,
      message: '尚有 ' + pendingRegisterOrders.length + ' 件商品未確定採購結果，若群組發送結帳訊息已超過24小時仍無法結帳，請主動聯繫 小幫手 或 W.W. 幫您確認！'
    };
  }

  if (batchIndices.length === 0) {
    return { success: false, message: '目前沒有待出貨結帳的連線商品！' };
  }

  // 計算運費 (6 大配送方式規則，動態連動系統設定表)
  const settings = getSystemSettings().data || {};
  const freeThreshold = Number(settings.FREE_SHIPPING_THRESHOLD) || 3000;
  const fee711 = Number(settings.SHIP_FEE_711_COD) || 38;
  const feeFami = Number(settings.SHIP_FEE_FAMI_COD) || 35;
  const feePostPrepaid = Number(settings.SHIP_FEE_POST_PREPAID) || 80;
  const feePostCod = Number(settings.SHIP_FEE_POST_COD) || 130;
  const feeBlackcat = Number(settings.SHIP_FEE_BLACKCAT) || 100;

  const METHOD_NAMES = {
    '711_PREPAID': '7-11 超商純取貨',
    'FAMI_PREPAID': '全家 超商純取貨',
    '711_COD': '7-11 超商取貨付款',
    'FAMI_COD': '全家 超商取貨付款',
    'POST_PREPAID': '郵局純寄件',
    'POST_COD': '郵局貨到付款',
    'BLACKCAT_PREPAID': '黑貓宅配到府',
    'STORE_PICKUP_FREE': '門市自取'
  };

  const method = checkoutData.shippingMethod; // 代碼
  const methodName = checkoutData.shippingMethodName || METHOD_NAMES[method] || method;
  let shippingFee = 0;
  let isPrepay = false; // 是否需先匯款

  if (method === '711_PREPAID' || method === 'FAMI_PREPAID' || method === 'STORE_PICKUP_FREE') {
    // 1 & 2. 7-11 或全家 超商純取貨 (需先匯款)：提供免運費 (NT$ 0)
    shippingFee = 0;
    isPrepay = true;
  } else if (method === '711_COD') {
    // 3. 7-11超商貨到付款 (賣貨便)：滿額免運
    shippingFee = (goodsTotal >= freeThreshold) ? 0 : fee711;
    isPrepay = false;
  } else if (method === 'FAMI_COD') {
    // 4. 全家超商貨到付款 (好賣+)：滿額免運
    shippingFee = (goodsTotal >= freeThreshold) ? 0 : feeFami;
    isPrepay = false;
  } else if (method === 'POST_PREPAID') {
    // 5. 郵局純寄件 (先匯款)
    shippingFee = feePostPrepaid;
    isPrepay = true;
  } else if (method === 'POST_COD') {
    // 6. 郵局貨到付款
    shippingFee = feePostCod;
    isPrepay = false;
  } else if (method === 'BLACKCAT_PREPAID') {
    // 7. 黑貓宅配到府 (先匯款)
    shippingFee = feeBlackcat;
    isPrepay = true;
  } else {
    shippingFee = Number(checkoutData.shippingFee) || 0;
  }

  const finalTotalAmount = goodsTotal + shippingFee;
  const payStatus = isPrepay ? '待付款' : '貨到付款待出貨';
  const orderStatus = '已完成結帳';
  const deliveryInfo = `[${methodName}] ${checkoutData.recipientAddress || ''}`;

  // 1. 更新所有採購成功列：第一筆記單筆運費，其餘記 0，寫入同一結帳編號 (第22欄)
  for (let idx = 0; idx < successfulIndices.length; idx++) {
    const r = successfulIndices[idx];
    const sub = Number(orderSheet.getRange(r, 10).getValue()) || 0;
    const fee = (idx === 0) ? shippingFee : 0; // 只有第一張帶運費
    const tot = sub + fee;

    orderSheet.getRange(r, 11).setValue(fee); // 運費
    orderSheet.getRange(r, 12).setValue(tot); // 總額
    orderSheet.getRange(r, 13).setValue(checkoutData.recipientName || '');
    orderSheet.getRange(r, 14).setValue(formatPhoneAsText(checkoutData.recipientPhone));
    orderSheet.getRange(r, 15).setValue(deliveryInfo);
    if (checkoutData.note) {
      orderSheet.getRange(r, 16).setValue(checkoutData.note);
    }
    
    // 若填寫了匯款後五碼，直接記錄並切換為已回報待對帳
    if (checkoutData.lastFive) {
      orderSheet.getRange(r, 17).setValue('已回報待對帳');
      orderSheet.getRange(r, 18).setValue("'" + String(checkoutData.lastFive).trim());
    } else {
      orderSheet.getRange(r, 17).setValue(payStatus);
    }
    orderSheet.getRange(r, 19).setValue(orderStatus);
    orderSheet.getRange(r, 22).setValue(checkoutId); // 第 22 欄：結帳編號
  }

  // 2. 更新缺貨/斷貨/取消列：同樣綁定同一結帳編號，運費為0，總額為0，標記免付款，結算時排除
  for (let i = 0; i < batchIndices.length; i++) {
    const r = batchIndices[i];
    if (successfulIndices.includes(r)) continue;

    orderSheet.getRange(r, 11).setValue(0); // 運費 0
    orderSheet.getRange(r, 12).setValue(0); // 應付總額 0 (缺貨免付)
    orderSheet.getRange(r, 13).setValue(checkoutData.recipientName || '');
    orderSheet.getRange(r, 14).setValue(formatPhoneAsText(checkoutData.recipientPhone));
    orderSheet.getRange(r, 15).setValue(deliveryInfo);
    orderSheet.getRange(r, 17).setValue('免付款(缺貨取消)');
    orderSheet.getRange(r, 19).setValue('缺貨斷貨');
    orderSheet.getRange(r, 22).setValue(checkoutId); // 第 22 欄：同樣綁定本批結帳編號！
  }

  // 更新顧客歸戶 (姓名與電話，不另外紀錄常用寄送地址)
  if (custSheet) {
    updateCustomerAddress(custSheet, userId, checkoutData.recipientName, checkoutData.recipientPhone);
  }

  return {
    success: true,
    message: '出貨結帳資料已確認送出！',
    checkoutId: checkoutId,
    batchCount: batchIndices.length,
    successfulCount: successfulIndices.length,
    goodsTotal: goodsTotal,
    shippingFee: shippingFee,
    totalAmount: finalTotalAmount,
    isPrepay: isPrepay,
    shippingMethod: method,
    shippingMethodName: checkoutData.shippingMethodName
  };
}

/**
 * 更新或建立顧客檔案歸戶
 * 欄位: ['LINE_User_ID', '最新暱稱', '真實姓名', '電話', '常用寄送地址', '歷史訂單數', '總消費金額', '首購日期', '最後下單日期', '黑名單標記']
 */
function updateCustomerProfile(custSheet, orderData, totalAmount, now) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!custSheet) {
      custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
      if (!custSheet) {
        custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
        custSheet.appendRow([
          'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '常用寄送地址', 
          '歷史訂單數', '總消費金額', '首購日期', '最後下單日期', '黑名單標記'
        ]);
        custSheet.getRange(1, 1, 1, 10).setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');
      }
    }

    const userId = orderData.userId || 'LINE_GUEST';
    const userName = orderData.userName || '訪客';
    const realName = orderData.realName || orderData.recipientName || '';
    const phone = formatPhoneAsText(orderData.phone);
    const data = custSheet.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        found = true;
        const row = i + 1;
        if (userName) custSheet.getRange(row, 2).setValue(userName);
        if (realName) custSheet.getRange(row, 3).setValue(realName);
        if (phone) custSheet.getRange(row, 4).setValue(phone);
        
        const currentOrders = Number(data[i][5]) || 0;
        const currentSpend = Number(data[i][6]) || 0;
        custSheet.getRange(row, 6).setValue(currentOrders + 1);
        custSheet.getRange(row, 7).setValue(currentSpend + (Number(totalAmount) || 0));
        if (!data[i][7]) {
          custSheet.getRange(row, 8).setValue(now);
        }
        custSheet.getRange(row, 9).setValue(now);
        break;
      }
    }

    if (!found) {
      custSheet.appendRow([
        userId,
        userName,
        realName,
        phone,
        '',
        1,
        Number(totalAmount) || 0,
        now,
        now,
        ''
      ]);
    }
  } catch (err) {
    console.error('更新顧客歸戶失敗 (不影響訂單建立): ' + err.toString());
  }
}

/**
 * 更新顧客歸戶姓名電話 (純文字格式防止 0 被吃掉，不另外紀錄常用寄送地址)
 */
function updateCustomerAddress(custSheet, userId, name, phone) {
  try {
    if (!custSheet) return;
    const cleanPhone = formatPhoneAsText(phone);
    const data = custSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const row = i + 1;
        if (name) custSheet.getRange(row, 3).setValue(name);
        if (cleanPhone) custSheet.getRange(row, 4).setValue(cleanPhone);
        // 不另外紀錄常用寄送地址
        break;
      }
    }
  } catch (err) {
    console.error('更新顧客歸戶失敗: ' + err.toString());
  }
}

/**
 * 格式化電話號碼為 Google 試算表純文字 (防止開頭 0 被自動轉為數字截斷)
 */
function formatPhoneAsText(phone) {
  if (phone === null || phone === undefined || phone === '') return '';
  const clean = String(phone).trim();
  if (!clean) return '';
  if (clean.startsWith("'")) return clean;
  return "'" + clean;
}

/**
 * 3. 客人回報匯款資料（支援匯款方式：LinePay、街口支付、中國信託轉帳、轉帳時間與後五碼）
 * 送出後將狀態更新為「對帳中，待出貨」
 */
function handleReportPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  let updatedCount = 0;

  const payMethod = data.payMethod || '中國信託轉帳';
  const transferTime = data.transferTime || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  const lastFive = data.lastFiveDigits ? "'" + String(data.lastFiveDigits).trim() : '';

  for (let i = 1; i < rows.length; i++) {
    const shipStatus = String(rows[i][18] || '').trim();
    // 嚴格排除缺貨斷貨、已取消商品
    if (shipStatus.includes('缺貨') || shipStatus.includes('斷貨') || shipStatus.includes('取消')) continue;

    const isTarget = (data.orderId && rows[i][0] === data.orderId) || 
                     (data.userId && rows[i][2] === data.userId && (rows[i][16] === '待付款' || rows[i][16] === '已回報待對帳' || shipStatus === '已完成結帳待出貨'));
    if (isTarget) {
      const rowIndex = i + 1;
      orderSheet.getRange(rowIndex, 17).setValue('對帳中，待出貨'); // 付款狀態
      if (lastFive) {
        orderSheet.getRange(rowIndex, 18).setValue(lastFive);       // 匯款後五碼
      }
      orderSheet.getRange(rowIndex, 19).setValue('已完成結帳');     // 採購/出貨狀態維持已完成結帳 (對齊下拉選單)

      // 處理備註記錄：匯款方式、轉帳時間與備註
      const noteText = `[付款回報: ${payMethod}] 時間: ${transferTime}` + (data.note ? ` 備註: ${data.note}` : '');
      orderSheet.getRange(rowIndex, 20).setValue(noteText);

      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    return { 
      success: true, 
      message: `🎉 已成功送出匯款回報！方式：${payMethod}，狀態已更新為「對帳中，待出貨」，小幫手確認入帳後將立即安排出貨！` 
    };
  }

  return { success: false, message: '查無符合條件的待付款訂單' };
}

/**
 * 4. 買家確認完成取貨（將狀態改為「已完成取貨」，正式歸檔至歷史訂單）
 */
function handleConfirmReceived(data) {
  const userId = data.userId;
  const orderId = data.orderId;
  const checkoutId = data.checkoutId;
  if (!userId) return { success: false, message: '缺少買家 ID' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < rows.length; i++) {
    const rUserId = String(rows[i][2] || '').trim();
    if (rUserId !== userId) continue;

    const rOrderId = String(rows[i][0] || '').trim();
    const rCheckoutId = String(rows[i][21] || '').trim();
    const shipStatus = String(rows[i][18] || '').trim();

    // 匹配條件：指定 checkoutId，或指定 orderId，或該買家所有「已完成出貨」項目
    const isTarget = (checkoutId && rCheckoutId === checkoutId) || 
                     (orderId && rOrderId === orderId) || 
                     (!checkoutId && !orderId && (shipStatus.includes('已完成出貨') || shipStatus.includes('出貨') || shipStatus.includes('配送')));
    if (isTarget) {
      orderSheet.getRange(i + 1, 19).setValue('已完成取貨');
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    return { success: true, message: '🎉 感謝您的回報！訂單已確認取件完畢並封存至歷史紀錄！' };
  }

  return { success: false, message: '查無符合條件的配送中訂單' };
}

/**
 * 5. 查詢使用者的訂單清單（區分待出貨結帳/配送中商品與歷史訂單，附帶顧客資訊、系統設定、包裹追蹤碼與結帳單號）
 */
function getOrdersForUser(userId) {
  if (!userId) return { success: false, message: '缺少 userId' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  const pendingCheckoutOrders = []; // 待出貨結帳或配送中（連線中登記/採購成功/已完成出貨配送中）
  const completedOrders = [];       // 已完成取貨結案之歷史訂單

  function formatDeliveryText(raw) {
    if (!raw) return '';
    let text = String(raw).trim();
    text = text.replace(/\[711_COD\]/gi, '[7-11 超商取貨付款]');
    text = text.replace(/\[FAMI_COD\]/gi, '[全家 超商取貨付款]');
    text = text.replace(/\[POST_COD\]/gi, '[郵局貨到付款]');
    text = text.replace(/\[711_PREPAID\]/gi, '[7-11 超商純取貨]');
    text = text.replace(/\[FAMI_PREPAID\]/gi, '[全家 超商純取貨]');
    text = text.replace(/\[POST_PREPAID\]/gi, '[郵局純寄件]');
    text = text.replace(/\[BLACKCAT_PREPAID\]/gi, '[黑貓宅配到府]');
    return text;
  }

  // 預先找出所有「已完成取貨」或「已結案」的結帳單號 (checkoutId)，確保該包裹之缺貨與成功品項整批同進歷史
  const closedCheckoutIds = new Set();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === userId) {
      const sStatus = String(rows[i][18] || '').trim();
      const cid = String(rows[i][21] || '').trim();
      if (cid && (sStatus.includes('已完成取貨') || sStatus.includes('結案') || sStatus.includes('已取消'))) {
        closedCheckoutIds.add(cid);
      }
    }
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[2] === userId) {
      const item = {
        orderId: row[0],
        orderDate: row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : row[1],
        productId: row[4],
        productName: row[5],
        spec: row[6],
        quantity: row[7],
        unitPrice: row[8],
        subtotal: row[9],
        shippingFee: row[10],
        totalAmount: row[11],
        recipientName: row[12],
        phone: row[13],
        deliveryAddress: formatDeliveryText(row[14]),
        note: row[15],
        paymentStatus: row[16],
        lastFive: row[17],
        shippingStatus: row[18],
        processNote: row[19] || '',
        trackingNumber: String(row[20] || '').trim(), // 第 21 欄 U 欄：包裹追蹤編號
        checkoutId: String(row[21] || '').trim()       // 第 22 欄 V 欄：結帳編號 (Checkout_ID)
      };

      const shipStatus = String(row[18] || '').trim();
      // 只有「已完成取貨」、「結案」或「已取消」或同結帳單號已結案之項目，才進入歷史訂單！
      const isClosed = shipStatus.includes('已完成取貨') || 
                       shipStatus.includes('結案') || 
                       shipStatus.includes('已取消') ||
                       (Boolean(item.checkoutId) && closedCheckoutIds.has(item.checkoutId));

      if (!isClosed) {
        pendingCheckoutOrders.push(item);
      } else {
        completedOrders.push(item);
      }
    }
  }

  // 取得顧客檔案
  let customerProfile = { userId: userId, realName: '', phone: '', defaultAddress: '' };
  if (custSheet) {
    const custData = custSheet.getDataRange().getValues();
    for (let c = 1; c < custData.length; c++) {
      if (custData[c][0] === userId) {
        customerProfile.userName = custData[c][1];
        customerProfile.realName = custData[c][2];
        customerProfile.phone = custData[c][3];
        customerProfile.defaultAddress = custData[c][4];
        break;
      }
    }
  }

  // 取得系統設定
  const settings = getSystemSettings().data || {};

  return { 
    success: true, 
    data: {
      pendingCheckoutOrders: pendingCheckoutOrders,
      completedOrders: completedOrders.reverse(),
      customerProfile: customerProfile,
      settings: settings
    }
  };
}

/**
 * 首次填寫聯絡資料彈窗時，立即正式建立/更新顧客檔案歸戶至試算表
 */
function handleRegisterCustomer(data) {
  if (!data) return { success: false, message: '缺少資料' };
  const userId = data.userId || 'LINE_GUEST';
  const userName = data.userName || '訪客';
  const realName = (data.realName || '').trim();
  const rawPhone = data.phone ? String(data.phone).trim() : '';
  const phone = formatPhoneAsText(rawPhone);

  if (!realName || !rawPhone) {
    return { success: false, message: '真實姓名與電話為必填' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!custSheet) {
    custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
    custSheet.appendRow([
      'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '常用寄送地址', 
      '歷史訂單數', '總消費金額', '首購日期', '最後下單日期', '黑名單標記'
    ]);
    custSheet.getRange(1, 1, 1, 10).setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');
    custSheet.getRange("D:D").setNumberFormat('@');
  }

  const now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
  const values = custSheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId) {
      found = true;
      const row = i + 1;
      if (userName) custSheet.getRange(row, 2).setValue(userName);
      custSheet.getRange(row, 3).setValue(realName);
      custSheet.getRange(row, 4).setValue(phone);
      custSheet.getRange(row, 9).setValue(now);
      break;
    }
  }

  if (!found) {
    custSheet.appendRow([
      userId,
      userName,
      realName,
      phone,
      '',
      0,
      0,
      '',
      now,
      ''
    ]);
  }

  return {
    success: true,
    message: '✅ 顧客資料已成功建立歸戶！',
    profile: {
      userId: userId,
      userName: userName,
      realName: realName,
      phone: rawPhone
    }
  };
}

/**
 * 取得系統設定
 */
function getSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return { success: false, message: '設定表不存在' };

  const rows = sheet.getDataRange().getValues();
  const settings = {};

  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0];
    const val = rows[i][1];
    if (key) {
      if (key === 'ORDER_DEADLINE' && val) {
        settings[key] = formatDeadlineStr(val);
      } else {
        settings[key] = val;
      }
    }
  }

  // 向下相容相容性處理：若舊試算表只有單一 BANK_INFO，自動拆解為銀行名、帳號與戶名
  if (settings.BANK_INFO && (!settings.BANK_NAME || !settings.BANK_ACCOUNT)) {
    const info = String(settings.BANK_INFO);
    if (!settings.BANK_NAME) {
      const bankMatch = info.match(/^([^\s]+(?:\s+[^\s]+)?)/);
      settings.BANK_NAME = bankMatch ? bankMatch[1] : '822 中國信託商業銀行';
    }
    if (!settings.BANK_ACCOUNT) {
      const accMatch = info.match(/帳號[:：\s]*([0-9-]+)/);
      settings.BANK_ACCOUNT = accMatch ? accMatch[1] : '4175-402-49273';
    }
    if (!settings.BANK_ACCOUNT_HOLDER) {
      const holderMatch = info.match(/戶名[:：\s]*([^\s]+)/);
      settings.BANK_ACCOUNT_HOLDER = holderMatch ? holderMatch[1] : '陳O雯';
    }
  }

  // 檢查試算表既有設定列並自動校正回寫 Google 試算表 (並刪除舊版 BANK_INFO 列)
  try {
    for (let i = rows.length - 1; i >= 1; i--) {
      const key = String(rows[i][0] || '').trim();
      const val = String(rows[i][1] || '').trim();
      if (key === 'BANK_INFO') {
        sheet.deleteRow(i + 1);
      } else if (key === 'BANK_ACCOUNT_HOLDER' && (val === '陳小美' || !val)) {
        sheet.getRange(i + 1, 2).setValue('陳O雯');
        settings.BANK_ACCOUNT_HOLDER = '陳O雯';
      } else if (key === 'BANK_ACCOUNT' && (!val || val.includes('123-4567') || val.includes('0123-4567') || val.includes('12345-6789'))) {
        sheet.getRange(i + 1, 2).setValue("'4175-402-49273");
        settings.BANK_ACCOUNT = '4175-402-49273';
      }
    }
  } catch (err) {
    console.warn('自動校正試算表失敗(可能無寫入權限):', err);
  }

  // 3種付款方式預設值保障與舊測試假帳號自動校正 (中國信託、LINE Pay、街口支付)
  if (!settings.BANK_NAME) settings.BANK_NAME = '822 中國信託商業銀行';
  if (!settings.BANK_ACCOUNT_HOLDER || settings.BANK_ACCOUNT_HOLDER === '陳小美') settings.BANK_ACCOUNT_HOLDER = '陳O雯';
  if (!settings.BANK_ACCOUNT || 
      settings.BANK_ACCOUNT.includes('123-4567') || 
      settings.BANK_ACCOUNT.includes('0123-4567') || 
      settings.BANK_ACCOUNT.includes('12345-6789')) {
    settings.BANK_ACCOUNT = '4175-402-49273';
  }
  if (!settings.LINEPAY_NAME) settings.LINEPAY_NAME = '粉絲團小編兼職小小幫手';
  if (!settings.LINEPAY_ACCOUNT) settings.LINEPAY_ACCOUNT = 'https://line.me/ti/p/IUqKCB7dRU';
  if (!settings.JKOPAY_NAME) settings.JKOPAY_NAME = 'Wen Wen Chen';
  if (!settings.JKOPAY_ACCOUNT) settings.JKOPAY_ACCOUNT = '900459491';

  // 系統規則校正：超商純取貨 0 元免運，貨到付款滿 3,000 元免運
  if (!settings.FREE_SHIPPING_THRESHOLD || Number(settings.FREE_SHIPPING_THRESHOLD) < 3000) {
    settings.FREE_SHIPPING_THRESHOLD = '3000';
  }
  settings.SHIP_FEE_STORE_PREPAID = '0'; // 超商純取貨一律免運
  if (!settings.SHIP_FEE_711_COD) settings.SHIP_FEE_711_COD = '38';
  if (!settings.SHIP_FEE_FAMI_COD) settings.SHIP_FEE_FAMI_COD = '35';

  return { success: true, data: settings };
}

/**
 * 專門一鍵更新 Google 試算表【系統設定】中的 3 大收款帳戶資訊
 * 可直接於 Apps Script 編輯器上方選擇此函式並點擊「執行」
 */
function updatePaymentSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settSheet) return '系統設定表不存在';

  const paymentConfigs = {
    'BANK_NAME': '822 中國信託商業銀行',
    'BANK_ACCOUNT_HOLDER': '陳O雯',
    'BANK_ACCOUNT': "'4175-402-49273",
    'LINEPAY_NAME': '粉絲團小編兼職小小幫手',
    'LINEPAY_ACCOUNT': 'https://line.me/ti/p/IUqKCB7dRU',
    'JKOPAY_NAME': 'Wen Wen Chen',
    'JKOPAY_ACCOUNT': "'900459491"
  };

  const data = settSheet.getDataRange().getValues();
  const updatedKeys = new Set();

  for (let r = 1; r < data.length; r++) {
    const key = String(data[r][0] || '').trim();
    if (paymentConfigs[key]) {
      settSheet.getRange(r + 1, 2).setValue(paymentConfigs[key]);
      updatedKeys.add(key);
    }
  }

  // 若試算表中尚無該欄位則自動 append 補齊
  const descriptions = {
    'BANK_NAME': '賣家收款銀行與代碼 (顯示於匯款卡片)',
    'BANK_ACCOUNT_HOLDER': '賣家帳戶戶名',
    'BANK_ACCOUNT': '賣家匯款帳號 (支援一鍵複製，以單引號確保純文字)',
    'LINEPAY_NAME': 'LINE Pay 帳戶顯示名稱',
    'LINEPAY_ACCOUNT': 'LINE Pay 加好友轉帳連結 (點擊直達或掃碼)',
    'JKOPAY_NAME': '街口支付帳戶顯示名稱',
    'JKOPAY_ACCOUNT': '街口支付帳號 (機構代碼396，支援一鍵複製)'
  };

  for (const [key, val] of Object.entries(paymentConfigs)) {
    if (!updatedKeys.has(key)) {
      settSheet.appendRow([key, val, descriptions[key] || '']);
    }
  }

  // 自動檢查並移除舊版單一欄位 BANK_INFO 列
  for (let r = data.length - 1; r >= 1; r--) {
    const key = String(data[r][0] || '').trim();
    if (key === 'BANK_INFO') {
      settSheet.deleteRow(r + 1);
    }
  }

  settSheet.getRange("B:B").setNumberFormat('@');
  return 'Google Sheets 系統設定資料已成功更新，且舊版 BANK_INFO 列已自動移除！';
}

/**
 * 自動產生商品編號：
 * 日幣商品: J + 西元年後2碼 + 日期(MMdd) + 0001 起流水號 (如 J2609270001)
 * 韓元商品: K + 西元年後2碼 + 日期(MMdd) + 0001 起流水號 (如 K2609270001)
 */
function generateProductId(sheet, currency, now) {
  const cur = String(currency || 'JPY').toUpperCase();
  const prefixChar = (cur === 'KRW' || cur.includes('韓') || cur.includes('KOR')) ? 'K' : 'J';
  const datePart = Utilities.formatDate(now || new Date(), 'Asia/Taipei', 'yyMMdd');
  const basePrefix = prefixChar + datePart; // 例如 J260927 或 K260927

  let maxSeq = 0;
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const existingId = String(data[i][0] || '').trim();
      if (existingId.startsWith(basePrefix)) {
        const suffix = existingId.substring(basePrefix.length);
        const seqNum = parseInt(suffix, 10);
        if (!isNaN(seqNum) && seqNum > maxSeq) {
          maxSeq = seqNum;
        }
      }
    }
  }

  const nextSeq = maxSeq + 1;
  let seqStr = String(nextSeq);
  while (seqStr.length < 4) {
    seqStr = '0' + seqStr;
  }
  return basePrefix + seqStr;
}

/**
 * 賣家快速新增商品 (可由賣家後台網頁調用，使用 LockService 避免連續建立併發衝突)
 */
function handleAddProduct(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: '系統忙碌中，請稍候重試 (Lock timeout)' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    if (!sheet) return { success: false, message: '商品表不存在' };

    if (!data.name || !data.price) {
      return { success: false, message: '商品名稱與價格為必填欄位！' };
    }

    const now = new Date();

    // 判斷幣別 (優先取 data.currency，或從 costNote / name / category 判斷)
    let cur = (data.currency || '').toUpperCase();
    if (!cur) {
      const hint = (data.name + ' ' + (data.costNote || '') + ' ' + (data.category || '')).toUpperCase();
      if (hint.includes('KRW') || hint.includes('韓') || hint.includes('KOR')) {
        cur = 'KRW';
      } else {
        cur = 'JPY';
      }
    }

    // 商品編號：若賣家有自訂則使用自訂，否則依「J/K + 西元年後2碼 + 日期 + 0001」自動跳號
    const pid = (data.id && String(data.id).trim()) || generateProductId(sheet, cur, now);

    // 規格處理 (支援陣列或逗號字串)
    let specsArray = [];
    if (Array.isArray(data.specs)) {
      specsArray = data.specs;
    } else if (typeof data.specs === 'string' && data.specs.trim()) {
      specsArray = data.specs.split(/[,，\n]/).map(s => s.trim()).filter(Boolean);
    }

    // 圖片處理：支援多張圖片（第一張為主圖）
    let imagesArray = [];
    if (Array.isArray(data.imageUrls) && data.imageUrls.length > 0) {
      imagesArray = data.imageUrls;
    } else if (data.imageUrl) {
      imagesArray = data.imageUrl.split(/[\n,]/).map(u => u.trim()).filter(Boolean);
    }
    const mainImage = imagesArray[0] || data.imageUrl || '';

    // 庫存名額：若勾選「不限庫存」則設定為 999999
    const stockQty = data.isUnlimitedStock ? 999999 : (Number(data.stock) || 10);

    sheet.appendRow([
      pid,
      data.name.trim(),
      data.category || '連線好物',
      (data.originalPrice && Number(data.originalPrice) > 0) ? Number(data.originalPrice) : '',
      Number(data.price),
      stockQty,
      JSON.stringify(specsArray),
      mainImage,
      data.description || '',
      '上架中',
      now,
      Number(data.costPrice) || 0, // 欄位 12: 成本價 (NT$)
      data.costNote || '',         // 欄位 13: 採購原幣與重量備註
      JSON.stringify(imagesArray), // 欄位 14: 所有圖片清單
      data.deadline ? String(data.deadline).trim() : '' // 欄位 15: 指定截單時間
    ]);

    SpreadsheetApp.flush();

    return { 
      success: true, 
      message: '商品建檔成功！', 
      productId: pid,
      product: {
        id: pid,
        name: data.name.trim(),
        category: data.category || '連線好物',
        originalPrice: (data.originalPrice && Number(data.originalPrice) > 0) ? Number(data.originalPrice) : 0,
        price: Number(data.price),
        stock: stockQty,
        specs: specsArray,
        imageUrl: mainImage,
        imageUrls: imagesArray,
        description: data.description || '',
        status: '上架中',
        costPrice: Number(data.costPrice) || 0,
        costNote: data.costNote || '',
        deadline: data.deadline ? String(data.deadline).trim() : ''
      }
    };
  } catch (err) {
    return { success: false, message: '商品建檔失敗：' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 處理附件圖片上傳至 Google Drive (極速優化版：利用資料夾繼承權限，省去逐檔遠端授權延遲)
 */
function handleUploadImages(data) {
  try {
    let folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch (e) {
      const folderName = 'LINE代購_商品圖片庫';
      const folders = DriveApp.getFoldersByName(folderName);
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const uploadedUrls = [];
    const files = data.files || [];

    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      let base64String = f.base64 || '';
      if (base64String.indexOf(',') > -1) {
        base64String = base64String.split(',')[1];
      }
      const decoded = Utilities.base64Decode(base64String);
      const contentType = f.type || 'image/jpeg';
      const fileName = 'prod_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss') + '_' + Math.floor(Math.random() * 10000) + '_' + (i + 1) + '.jpg';
      
      const blob = Utilities.newBlob(decoded, contentType, fileName);
      const file = folder.createFile(blob);
      // 💡 資料夾已設為公開檢視，檔案建立時自動繼承公開權限，無需逐檔 setSharing，每張照片節省 1~1.5 秒！

      // 直接輸出 Google Drive 穩定可直連的圖片網址
      const directUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      uploadedUrls.push(directUrl);
    }

    return { success: true, urls: uploadedUrls };
  } catch (err) {
    return { success: false, message: '圖片上傳至 Google Drive 失敗：' + err.toString() };
  }
}

/**
 * 測試並激活 Google Drive 權限（可在 GAS 編輯器上方下拉選單直接點「執行」測試）
 */
function testDrivePermission() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  Logger.log('成功連接 Google Drive 資料夾：' + folder.getName());
  return 'OK: ' + folder.getName();
}

/**
 * 工具函式：輸出 JSON
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🛠️ 清除測試資料管理函式
 * 說明：清空「訂單明細」與「顧客歸戶」的所有測試資料，保留第一列表頭標題與欄位格式設定
 * 可在 Google Apps Script 編輯器中直接選取此函式點擊「執行」，亦可透過 API 觸發
 */
function clearTestOrdersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deletedOrdersCount = 0;
  let deletedCustCount = 0;

  // 1. 清除「訂單明細」測試訂單 (保留第 1 列表頭)
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (orderSheet) {
    const lastRow = orderSheet.getLastRow();
    if (lastRow > 1) {
      deletedOrdersCount = lastRow - 1;
      orderSheet.getRange(2, 1, lastRow - 1, orderSheet.getLastColumn()).clearContent();
    }
  }

  // 2. 清除「顧客歸戶」測試資料 (保留第 1 列表頭)
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (custSheet) {
    const lastRow = custSheet.getLastRow();
    if (lastRow > 1) {
      deletedCustCount = lastRow - 1;
      custSheet.getRange(2, 1, lastRow - 1, custSheet.getLastColumn()).clearContent();
    }
  }

  Logger.log(`✅ 清除完成！共清除 ${deletedOrdersCount} 筆訂單明細與 ${deletedCustCount} 筆顧客資料。已保留表頭與格式設定。`);
  return {
    success: true,
    message: `✅ 測試資料已成功清除完畢！共清除 ${deletedOrdersCount} 筆訂單與 ${deletedCustCount} 筆顧客歸戶資料，已保留表頭結構與格式。`,
    deletedOrders: deletedOrdersCount,
    deletedCustomers: deletedCustCount
  };
}

/**
 * 🛠️ 清除自動化測試商品 (不影響任何正式商品)
 * 清除名稱包含「測試連續新增商品」或「自動連續測試商品」或「即時測試商品」的商品資料
 */
function clearTestProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  if (!sheet) return { success: false, message: '商品表不存在' };

  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const pName = String(data[i][1] || '');
    const pCat = String(data[i][2] || '');
    if (pName.includes('測試') || pCat === '測試' || pName.startsWith('自動連續測試') || pName.startsWith('即時測試') || pName.startsWith('商品1') || pName.startsWith('商品2') || pName.startsWith('商品3') || pName.startsWith('商品4') || pName.startsWith('商品5')) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return {
    success: true,
    message: `✅ 已清除 ${deletedCount} 筆測試商品！`,
    deletedCount: deletedCount
  };
}

/**
 * 處理頁面與商品瀏覽人次累計 (使用 LockService 避免併發衝突)
 */
function handleRecordPageView(data) {
  if (!data || !data.page) {
    return { success: false, message: '無效的頁面資料' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: '伺服器繁忙，略過本次統計' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.VIEWS || '瀏覽統計');
    const headers = ['頁面名稱', '頁面路徑/識別碼', '商品編號', '累計瀏覽次數', '最後瀏覽時間'];

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.VIEWS || '瀏覽統計');
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#4f46e5').setFontColor('#ffffff').setFontWeight('bold');
      sheet.getRange("D:D").setNumberFormat('#,##0');
      sheet.getRange("E:E").setNumberFormat('yyyy/MM/dd HH:mm:ss');
    }

    const pagePath = String(data.page || '').trim();
    const pageTitle = String(data.title || pagePath).trim();
    const productId = String(data.productId || '-').trim();

    // 嚴格排除「賣家管理與發卡推播中心」與「買家訂購清單與合併結帳」，不計入公開瀏覽量
    const pLower = pagePath.toLowerCase();
    if (pLower === 'admin-card-generator.html' || 
        pLower === 'my-orders.html' || 
        pageTitle.includes('賣家管理') || 
        pageTitle.includes('買家訂購清單') || 
        pageTitle.includes('合併結帳')) {
      return { success: true, ignored: true, message: '管理中心與訂單頁不計入公開瀏覽人次' };
    }

    const now = new Date();

    const range = sheet.getDataRange();
    const values = range.getValues();
    let targetRow = -1;
    let currentViews = 0;

    for (let r = 1; r < values.length; r++) {
      const existingPath = String(values[r][1] || '').trim();
      const existingPid = String(values[r][2] || '').trim();
      
      if (productId !== '-' && productId !== '') {
        if (existingPid === productId || existingPath === pagePath) {
          targetRow = r + 1;
          currentViews = Number(values[r][3]) || 0;
          break;
        }
      } else {
        if (existingPath === pagePath) {
          targetRow = r + 1;
          currentViews = Number(values[r][3]) || 0;
          break;
        }
      }
    }

    const newViews = currentViews + 1;

    if (targetRow > 0) {
      if (pageTitle) sheet.getRange(targetRow, 1).setValue(pageTitle);
      if (productId && productId !== '-') sheet.getRange(targetRow, 3).setValue(productId);
      sheet.getRange(targetRow, 4).setValue(newViews);
      sheet.getRange(targetRow, 5).setValue(now);
    } else {
      sheet.appendRow([pageTitle, pagePath, productId, 1, now]);
    }

    return { success: true, page: pagePath, views: newViews };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 取得全站各頁面與商品瀏覽統計數據 (嚴格排除管理中心與訂單頁)
 */
function getViewStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.VIEWS || '瀏覽統計');
  if (!sheet) {
    return { success: true, totalViews: 0, pages: [], topProducts: [] };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { success: true, totalViews: 0, pages: [], topProducts: [] };
  }

  // 判斷是否為應排除的內部/個人頁面
  function isExcludedPageView(path, title) {
    const p = String(path || '').toLowerCase();
    const t = String(title || '');
    return p === 'admin-card-generator.html' || 
           p === 'my-orders.html' || 
           t.includes('賣家管理') || 
           t.includes('買家訂購清單') || 
           t.includes('合併結帳');
  }

  // 自動清理試算表中既有的管理中心與訂單頁面列 (從最後一列往上刪除避免行號偏移)
  try {
    for (let r = values.length - 1; r >= 1; r--) {
      const title = String(values[r][0] || '').trim();
      const path = String(values[r][1] || '').trim();
      if (isExcludedPageView(path, title)) {
        sheet.deleteRow(r + 1);
      }
    }
  } catch (cleanErr) {
    console.warn('清理排除頁面列失敗:', cleanErr);
  }

  const currentValues = sheet.getDataRange().getValues();

  let totalViews = 0;
  const pages = [];
  const productViewsMap = {};

  for (let r = 1; r < currentValues.length; r++) {
    const title = String(currentValues[r][0] || '').trim();
    const path = String(currentValues[r][1] || '').trim();
    const pid = String(currentValues[r][2] || '').trim();
    const views = Number(currentValues[r][3]) || 0;
    const lastTime = currentValues[r][4] ? Utilities.formatDate(new Date(currentValues[r][4]), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : '';

    if (isExcludedPageView(path, title)) continue;

    totalViews += views;

    const item = {
      title: title,
      path: path,
      productId: pid,
      views: views,
      lastVisited: lastTime
    };

    pages.push(item);

    if (pid && pid !== '-') {
      productViewsMap[pid] = (productViewsMap[pid] || 0) + views;
    }
  }

  pages.sort((a, b) => b.views - a.views);

  // 取得商品清單進行資訊豐富化 (圖片、價格)
  const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  const productsMeta = {};
  if (prodSheet) {
    const prodData = prodSheet.getDataRange().getValues();
    for (let r = 1; r < prodData.length; r++) {
      const pid = String(prodData[r][0] || '').trim();
      const pName = String(prodData[r][1] || '').trim();
      const pImg = String(prodData[r][7] || '').trim();
      const pPrice = Number(prodData[r][4]) || 0;
      if (pid) {
        productsMeta[pid] = { name: pName, imageUrl: pImg, price: pPrice };
      }
    }
  }

  const topProducts = [];
  for (const pid in productViewsMap) {
    const meta = productsMeta[pid] || {};
    topProducts.push({
      productId: pid,
      name: meta.name || pid,
      imageUrl: meta.imageUrl || '',
      price: meta.price || 0,
      views: productViewsMap[pid]
    });
  }
  topProducts.sort((a, b) => b.views - a.views);

  return {
    success: true,
    totalViews: totalViews,
    pages: pages,
    topProducts: topProducts
  };
}


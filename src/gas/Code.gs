/**
 * LINE 連線代購 & 群組團購整合系統 - GAS 後端 API
 * 專為 Google Sheets + LINE LIFF 架構設計
 */

// 工作表名稱常數
const SHEET_NAMES = {
  PRODUCTS: '商品清單',
  ORDERS: '訂單明細',
  CUSTOMERS: '顧客歸戶',
  SETTINGS: '系統設定'
};

// 指定的商品圖片 Google Drive 資料夾 ID
const DRIVE_FOLDER_ID = '1dqoqS4VJK7Fn56-L-mM5Afwyb6dvTQMB';

/**
 * 試算表初次安裝設定與欄位升級：自動補齊成本價、重量備註與多圖欄位
 * 可在 Apps Script 編輯器中直接執行此函式升級試算表結構
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 商品清單工作表
  let prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  const prodHeaders = [
    '商品編號', '商品名稱', '分類', '專櫃原價', '連線代購價', '現貨庫存', 
    '規格清單(JSON)', '封面主圖網址', '商品描述', '狀態', '建立時間',
    '成本價(NT$)', '採購原幣與重量備註', '所有圖片清單(JSON)'
  ];

  if (!prodSheet) {
    prodSheet = ss.insertSheet(SHEET_NAMES.PRODUCTS);
    prodSheet.appendRow(prodHeaders);
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    
    // 加入範例資料
    prodSheet.appendRow([
      'P001', '日本代購限定 輕量防潑水後背包', '包包配件', 1880, 1450, 999999,
      JSON.stringify(['米白色', '經典黑', '海軍藍']),
      'https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80',
      '日本專櫃直購，輕盈大容量，防潑水尼龍材質！', '上架中', new Date(),
      850, 'JPY 3500 (含稅), 380g', JSON.stringify(['https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80'])
    ]);
  } else {
    // 既有表格：升級表頭為最新規格
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setValues([prodHeaders]);
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 2. 訂單明細工作表
  let orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) {
    orderSheet = ss.insertSheet(SHEET_NAMES.ORDERS);
    orderSheet.appendRow([
      '訂單編號', '下單時間', 'LINE_User_ID', 'LINE暱稱', '商品編號', 
      '商品名稱', '選購規格', '數量', '單價', '商品小計', '運費', 
      '訂單總額', '收件人姓名', '聯絡電話', '取件方式與地址', 
      '買家備註', '付款狀態', '匯款後五碼', '出貨狀態', '處理備註'
    ]);
    orderSheet.getRange(1, 1, 1, 20).setBackground('#107c41').setFontColor('#ffffff').setFontWeight('bold');
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

  // 4. 系統設定表
  let settSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settSheet) {
    settSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settSheet.appendRow(['設定項目', '設定值', '說明']);
    settSheet.getRange(1, 1, 1, 3).setBackground('#d83b01').setFontColor('#ffffff').setFontWeight('bold');
    settSheet.appendRow(['BANK_INFO', '822 中國信託 帳號: 12345-67890123 戶名: 代購小幫手', '賣家匯款帳號資料']);
    settSheet.appendRow(['DEFAULT_SHIPPING_FEE', '60', '預設超取運費']);
    settSheet.appendRow(['FREE_SHIPPING_THRESHOLD', '1500', '滿額免運門檻']);
    settSheet.appendRow(['STORE_NAME', '日韓嚴選連線代購', '商店名稱']);
  }

  return '工作表初始化完成！';
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
    } else {
      result = { success: false, message: '未知的 action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
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
    } else if (action === 'checkoutOrders') {
      result = handleCheckoutOrders(postData.data);
    } else if (action === 'reportPayment') {
      result = handleReportPayment(postData.data);
    } else if (action === 'addProduct') {
      result = handleAddProduct(postData.data);
    } else if (action === 'uploadImages') {
      result = handleUploadImages(postData.data);
    } else {
      result = { success: false, message: '未知的 action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
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
        costNote: row[12] || ''
      });
    }
  }

  return { success: true, data: products };
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

    // 1. 檢查商品庫存
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

    // 3. 產生訂單編號
    const now = new Date();
    const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss');
    const randomSuffix = Math.floor(100 + Math.random() * 900);
    const orderId = 'OD' + timeStr + randomSuffix;

    const unitPrice = Number(targetProduct[4]); // 連線代購價
    const subtotal = unitPrice * buyQty;
    const shippingFee = 0; // 連線期間運費先設為 0，出貨結帳時統一合併計算
    const totalAmount = subtotal;

    // 4. 寫入訂單明細（狀態為「連線中待出貨」，地址為待出貨填寫）
    orderSheet.appendRow([
      orderId,
      now,
      orderData.userId || 'LINE_GUEST',
      orderData.userName || '訪客',
      orderData.productId,
      targetProduct[1], // 品名
      orderData.spec || '單一規格',
      buyQty,
      unitPrice,
      subtotal,
      shippingFee,
      totalAmount,
      orderData.realName || orderData.recipientName || '',
      orderData.phone || '',
      '[待出貨結帳填寫]',
      orderData.note || '',
      '未結帳',
      '',
      '連線中待出貨', // 訂單處理狀態
      ''
    ]);

    // 5. 更新或建立顧客檔案歸戶
    if (orderData.realName || orderData.phone) {
      updateCustomerProfile(custSheet, orderData, totalAmount, now);
    }

    return {
      success: true,
      message: '🎉 搶購登記成功！已為您保留商品名額！',
      orderId: orderId,
      subtotal: subtotal,
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
  const pendingIndices = []; // 紀錄所有「連線中待出貨」或「未結帳」的列號
  let goodsTotal = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    // 比對 userId，且狀態為連線中待出貨或未結帳
    if (row[2] === userId && (row[18] === '連線中待出貨' || row[16] === '未結帳')) {
      pendingIndices.push(i + 1); // 1-based row index
      goodsTotal += (Number(row[9]) || 0); // 累計商品小計
    }
  }

  if (pendingIndices.length === 0) {
    return { success: false, message: '目前沒有待出貨結帳的連線商品！' };
  }

  // 計算運費 (6 大配送方式規則)
  const method = checkoutData.shippingMethod; // 代碼
  let shippingFee = 0;
  let isPrepay = false; // 是否需先匯款

  if (method === 'STORE_PICKUP_FREE') {
    // 1. 超商純取貨 (需先匯款)：不限購買金額，運費0元
    shippingFee = 0;
    isPrepay = true;
  } else if (method === '711_COD') {
    // 2. 7-11超商貨到付款 (賣貨便)：運費38元/滿3000元運費0元
    shippingFee = (goodsTotal >= 3000) ? 0 : 38;
    isPrepay = false;
  } else if (method === 'FAMI_COD') {
    // 3. 全家超商貨到付款 (好賣+)：運費35元/滿3000元運費0元
    shippingFee = (goodsTotal >= 3000) ? 0 : 35;
    isPrepay = false;
  } else if (method === 'POST_PREPAID') {
    // 4. 郵局純寄件 (先匯款)：郵資80元
    shippingFee = 80;
    isPrepay = true;
  } else if (method === 'POST_COD') {
    // 5. 郵局貨到付款：郵資130元
    shippingFee = 130;
    isPrepay = false;
  } else if (method === 'BLACKCAT_PREPAID') {
    // 6. 黑貓宅配到府 (先匯款)：運費100元
    shippingFee = 100;
    isPrepay = true;
  } else {
    shippingFee = Number(checkoutData.shippingFee) || 0;
  }

  const finalTotalAmount = goodsTotal + shippingFee;
  const payStatus = isPrepay ? '待付款' : '貨到付款待出貨';
  const orderStatus = '已完成結帳待出貨';
  const deliveryInfo = `[${checkoutData.shippingMethodName || method}] ${checkoutData.recipientAddress || ''}`;

  // 更新所有待結帳列：第一筆記單筆運費，其餘記 0，避免運費重複加總
  for (let idx = 0; idx < pendingIndices.length; idx++) {
    const r = pendingIndices[idx];
    const sub = Number(orderSheet.getRange(r, 10).getValue()) || 0;
    const fee = (idx === 0) ? shippingFee : 0; // 只有第一張帶運費
    const tot = sub + fee;

    orderSheet.getRange(r, 11).setValue(fee); // 運費
    orderSheet.getRange(r, 12).setValue(tot); // 總額
    orderSheet.getRange(r, 13).setValue(checkoutData.recipientName || '');
    orderSheet.getRange(r, 14).setValue(checkoutData.recipientPhone || '');
    orderSheet.getRange(r, 15).setValue(deliveryInfo);
    if (checkoutData.note) {
      orderSheet.getRange(r, 16).setValue(checkoutData.note);
    }
    orderSheet.getRange(r, 17).setValue(payStatus);
    orderSheet.getRange(r, 19).setValue(orderStatus);
  }

  // 更新顧客歸戶地址
  if (custSheet && checkoutData.recipientAddress) {
    updateCustomerAddress(custSheet, userId, checkoutData.recipientName, checkoutData.recipientPhone, checkoutData.recipientAddress);
  }

  return {
    success: true,
    message: '出貨結帳資料已確認送出！',
    orderCount: pendingIndices.length,
    goodsTotal: goodsTotal,
    shippingFee: shippingFee,
    totalAmount: finalTotalAmount,
    isPrepay: isPrepay,
    shippingMethod: method,
    shippingMethodName: checkoutData.shippingMethodName
  };
}

/**
 * 更新顧客常用地址電話
 */
function updateCustomerAddress(custSheet, userId, name, phone, address) {
  const data = custSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      const row = i + 1;
      if (name) custSheet.getRange(row, 3).setValue(name);
      if (phone) custSheet.getRange(row, 4).setValue(phone);
      if (address) custSheet.getRange(row, 5).setValue(address);
      break;
    }
  }
}

/**
 * 3. 客人回填匯款後五碼（支援依 userId 一次性更新所有待付款訂單）
 */
function handleReportPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < rows.length; i++) {
    const isTarget = (data.orderId && rows[i][0] === data.orderId) || 
                     (data.userId && rows[i][2] === data.userId && (rows[i][16] === '待付款' || rows[i][16] === '未結帳'));
    if (isTarget) {
      const rowIndex = i + 1;
      orderSheet.getRange(rowIndex, 17).setValue('已回報待對帳');
      orderSheet.getRange(rowIndex, 18).setValue("'" + String(data.lastFiveDigits));
      if (data.note) {
        orderSheet.getRange(rowIndex, 20).setValue('買家留言: ' + data.note);
      }
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    return { success: true, message: `已成功回報匯款後五碼 (${data.lastFiveDigits})！賣家核對入帳後將為您準備出貨包裹。` };
  }

  return { success: false, message: '查無符合條件的待付款訂單' };
}

/**
 * 4. 查詢使用者的訂單清單（區分待出貨結帳商品與歷史訂單，附帶顧客資訊與銀行帳戶）
 */
function getOrdersForUser(userId) {
  if (!userId) return { success: false, message: '缺少 userId' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  const pendingCheckoutOrders = []; // 待出貨結帳（連線中登記）
  const completedOrders = [];       // 已完成結帳或歷史訂單

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
        deliveryAddress: row[14],
        note: row[15],
        paymentStatus: row[16],
        lastFive: row[17],
        shippingStatus: row[18]
      };

      if (row[18] === '連線中待出貨' || row[16] === '未結帳') {
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
    if (key) settings[key] = val;
  }

  return { success: true, data: settings };
}

/**
 * 賣家快速新增商品 (可由賣家後台網頁調用)
 */
function handleAddProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  if (!sheet) return { success: false, message: '商品表不存在' };

  if (!data.name || !data.price) {
    return { success: false, message: '商品名稱與價格為必填欄位！' };
  }

  const now = new Date();
  const pid = (data.id && String(data.id).trim()) || ('P' + Utilities.formatDate(now, 'Asia/Taipei', 'MMddHHmm'));

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
    JSON.stringify(imagesArray)  // 欄位 14: 所有圖片清單
  ]);

  return { 
    success: true, 
    message: '商品建檔成功！', 
    productId: pid,
    product: {
      id: pid,
      name: data.name,
      price: Number(data.price),
      imageUrl: mainImage,
      costPrice: Number(data.costPrice) || 0
    }
  };
}

/**
 * 處理附件圖片上傳至 Google Drive
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
    }
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

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
      const fileName = 'prod_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss') + '_' + (i + 1) + '.jpg';
      
      const blob = Utilities.newBlob(decoded, contentType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

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

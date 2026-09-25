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

/**
 * 試算表初次安裝設定：自動建立所需的工作表與欄位標題
 * 可在 Apps Script 編輯器中直接執行此函式
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 商品清單工作表
  let prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  if (!prodSheet) {
    prodSheet = ss.insertSheet(SHEET_NAMES.PRODUCTS);
    prodSheet.appendRow([
      '商品編號', '商品名稱', '分類', '原價', '特價', '庫存', 
      '規格清單(JSON)', '圖片網址', '商品描述', '狀態', '建立時間'
    ]);
    // 預設樣式
    prodSheet.getRange(1, 1, 1, 11).setBackground('#2b5797').setFontColor('#ffffff').setFontWeight('bold');
    
    // 加入範例資料
    prodSheet.appendRow([
      'P001', '日本代購限定 輕量防潑水後背包', '包包配件', 1880, 1450, 15,
      JSON.stringify(['米白色', '經典黑', '海軍藍']),
      'https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80',
      '日本專櫃直購，輕盈大容量，防潑水尼龍材質！', '上架中', new Date()
    ]);
    prodSheet.appendRow([
      'P002', '熱銷款 舒眠天然草本香氛噴霧 100ml', '生活居家', 850, 680, 30,
      JSON.stringify(['薰衣草森林', '洋甘菊微風']),
      'https://images.unsplash.com/photo-1608571423902-eed4a5ad8108?w=800&auto=format&fit=crop&q=80',
      '日本飯店御用款，睡前噴在枕頭上放鬆助眠。', '上架中', new Date()
    ]);
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
 * - createOrder (下單寫入 + 庫存檢查防超賣)
 * - reportPayment (客人回填匯款後五碼)
 * - addProduct (賣家快速上架商品)
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    let result = {};
    if (action === 'createOrder') {
      result = handleCreateOrder(postData.data);
    } else if (action === 'reportPayment') {
      result = handleReportPayment(postData.data);
    } else if (action === 'addProduct') {
      result = handleAddProduct(postData.data);
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

      products.push({
        id: row[0],
        name: row[1],
        category: row[2],
        originalPrice: Number(row[3]),
        price: Number(row[4]),
        stock: Number(row[5]),
        specs: specs,
        imageUrl: row[7],
        description: row[8],
        status: status
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
 * 下單處理（含防超賣排隊鎖 LockService）
 */
function handleCreateOrder(orderData) {
  // 建立系統排隊鎖，防止同秒數搶購導致超賣
  const lock = LockService.getScriptLock();
  try {
    // 最多等待 10 秒鎖定
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
        productRowIndex = i + 1; // 1-based row index
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

    // 3. 產生訂單編號 (格式: OD + 年月日時分秒 + 隨機3碼)
    const now = new Date();
    const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss');
    const randomSuffix = Math.floor(100 + Math.random() * 900);
    const orderId = 'OD' + timeStr + randomSuffix;

    const unitPrice = Number(targetProduct[4]); // 特價
    const subtotal = unitPrice * buyQty;
    const shippingFee = Number(orderData.shippingFee || 60);
    const totalAmount = subtotal + shippingFee;

    // 4. 寫入訂單明細
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
      orderData.recipientName || '',
      orderData.phone || '',
      `[${orderData.shippingMethod || '超商取貨'}] ${orderData.deliveryAddress || ''}`,
      orderData.note || '',
      '待付款',
      '', // 匯款後五碼尚未回填
      '未出貨',
      ''
    ]);

    // 5. 更新或建立顧客檔案歸戶
    updateCustomerProfile(custSheet, orderData, totalAmount, now);

    return {
      success: true,
      message: '下單成功！',
      orderId: orderId,
      totalAmount: totalAmount,
      productName: targetProduct[1],
      remainingStock: newStock
    };

  } finally {
    // 釋放鎖
    lock.releaseLock();
  }
}

/**
 * 顧客檔案歸戶更新
 */
function updateCustomerProfile(custSheet, orderData, orderAmount, timestamp) {
  if (!custSheet || !orderData.userId) return;

  const data = custSheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderData.userId) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow !== -1) {
    // 更新既有客戶
    const prevOrders = Number(custSheet.getRange(foundRow, 6).getValue()) || 0;
    const prevSpent = Number(custSheet.getRange(foundRow, 7).getValue()) || 0;
    
    custSheet.getRange(foundRow, 2).setValue(orderData.userName || '');
    if (orderData.recipientName) custSheet.getRange(foundRow, 3).setValue(orderData.recipientName);
    if (orderData.phone) custSheet.getRange(foundRow, 4).setValue(orderData.phone);
    if (orderData.deliveryAddress) custSheet.getRange(foundRow, 5).setValue(orderData.deliveryAddress);
    custSheet.getRange(foundRow, 6).setValue(prevOrders + 1);
    custSheet.getRange(foundRow, 7).setValue(prevSpent + orderAmount);
    custSheet.getRange(foundRow, 9).setValue(timestamp);
  } else {
    // 新增顧客
    custSheet.appendRow([
      orderData.userId,
      orderData.userName || '',
      orderData.recipientName || '',
      orderData.phone || '',
      orderData.deliveryAddress || '',
      1,
      orderAmount,
      timestamp,
      timestamp,
      '正常'
    ]);
  }
}

/**
 * 客人回填匯款後五碼
 */
function handleReportPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.orderId) {
      const rowIndex = i + 1;
      // 填入後五碼與更新付款狀態為「已回報對帳中」
      orderSheet.getRange(rowIndex, 17).setValue('已回報待對帳');
      orderSheet.getRange(rowIndex, 18).setValue("'" + String(data.lastFiveDigits));
      if (data.note) {
        orderSheet.getRange(rowIndex, 20).setValue('買家留言: ' + data.note);
      }
      return { success: true, message: '回報成功！賣家將於核對入帳後為您安排代購出貨。' };
    }
  }

  return { success: false, message: '查無此訂單編號' };
}

/**
 * 查詢特定使用者的所有歷史訂單
 */
function getOrdersForUser(userId) {
  if (!userId) return { success: false, message: '缺少 userId' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) return { success: false, message: '訂單表不存在' };

  const rows = orderSheet.getDataRange().getValues();
  const userOrders = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[2] === userId) {
      userOrders.push({
        orderId: row[0],
        orderDate: row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : row[1],
        productId: row[4],
        productName: row[5],
        spec: row[6],
        quantity: row[7],
        subtotal: row[9],
        shippingFee: row[10],
        totalAmount: row[11],
        paymentStatus: row[16],
        lastFive: row[17],
        shippingStatus: row[18]
      });
    }
  }

  return { success: true, data: userOrders.reverse() }; // 最新訂單排在前面
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

  const now = new Date();
  const pid = 'P' + Utilities.formatDate(now, 'Asia/Taipei', 'MMddHHmm');

  sheet.appendRow([
    pid,
    data.name,
    data.category || '連線好物',
    Number(data.originalPrice) || Number(data.price),
    Number(data.price),
    Number(data.stock) || 10,
    JSON.stringify(data.specs || []),
    data.imageUrl || '',
    data.description || '',
    '上架中',
    now
  ]);

  return { success: true, message: '商品上架成功！', productId: pid };
}

/**
 * 工具函式：輸出 JSON
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ==============================================================================
 * 【04_CustomerService.gs】顧客歸戶與會員中心服務模組 (CRM & Profile)
 * ==============================================================================
 * 說明：專門處理顧客會員歸戶、收件個資自動維護、生命週期動態 Log 與 1對1結帳通報清單。
 * ==============================================================================
 */

/**
  * 檢查是否為已建檔之老顧客
  */
  function checkCustomerExists(userId) {
    if (!userId) return { success: false, exists: false };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    if (!custSheet) return { success: true, exists: false };

    // 確保 13 欄結構支援最近收件人姓名與電話
    ensureOrderAndCustomerSchema(ss, null, custSheet);

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
            joinedAt: data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], 'GMT+8', 'yyyy/MM/dd HH:mm:ss') : String(data[i][4] || ''),
            lastVisitAt: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], 'GMT+8', 'yyyy/MM/dd HH:mm:ss') : String(data[i][5] || ''),
            lastViewed: data[i][6] || '',
            visitDays: Number(data[i][7]) || 1,
            visitCount: Number(data[i][7]) || 1, // 相容舊版名稱
            pageViews: Number(data[i][8]) || 1,
            orderCount: Number(data[i][9]) || 0,
            totalSpend: Number(data[i][10]) || 0,
            memberTier: data[i][11] || '一般會員',
            lifeStatus: data[i][12] || '🟢 活躍',
            source: data[i][13] || '',
            defaultAddress: data[i][14] || '',
            isBlacklisted: data[i][15] || '正常',
            notes: data[i][16] || '',
            lastRecipientName: data[i][2] || '',
            lastRecipientPhone: data[i][3] || ''
          }
        };
      }
    }
    return { success: true, exists: false };
  }

/**
  * 首次填寫聯絡資料彈窗時，立即正式建立/更新顧客檔案歸戶至試算表
  */
  function handleRegisterCustomer(data) {
    if (!data) return { success: false, message: '缺少資料' };
    let userId = String(data.userId || '').trim();
    const userName = data.userName || 'LINE 會員';
    const realName = (data.realName || '').trim();
    const rawPhone = data.phone ? String(data.phone).trim() : '';
    const cleanPhone = rawPhone.replace(/[-\s]/g, '');
    const phone = formatPhoneAsText(rawPhone);
    const source = (data.source || '').trim();

    if (!realName || !rawPhone) {
      return { success: false, message: '真實姓名與電話為必填' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    if (!custSheet) {
      custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
      custSheet.appendRow([
        'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '加入會員時間', 
        '最後進站時間', '最後瀏覽商品/頁面', '進站累計天數', '商品與頁面瀏覽次數', '歷史訂單數', '總消費金額', 
        '會員等級', '顧客生命狀態', '顧客來源', '常用超商取件門市', '黑名單標記', '賣家管理備註/標籤'
      ]);
      custSheet.getRange(1, 1, 1, 17).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
      custSheet.getRange("D:D").setNumberFormat('@');
    }

    ensureOrderAndCustomerSchema(ss, null, custSheet);

    const now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
    const values = custSheet.getDataRange().getValues();
    let found = false;

    const isRealLineId = userId && !userId.startsWith('GUEST_') && userId !== 'TEST_USER_999' && userId !== 'LINE_GUEST';

    for (let i = 1; i < values.length; i++) {
      const rowUid = String(values[i][0] || '').trim();
      const rowPhone = String(values[i][3] || '').replace(/[-\s]/g, '').trim();

      const isMatched = (rowUid === userId) || (isRealLineId && rowPhone === cleanPhone && (rowUid.startsWith('GUEST_') || rowUid === 'TEST_USER_999'));

      if (isMatched) {
        found = true;
        const row = i + 1;
        if (isRealLineId && (rowUid.startsWith('GUEST_') || rowUid === 'TEST_USER_999')) {
          custSheet.getRange(row, 1).setValue(userId);
        }
        if (userName) custSheet.getRange(row, 2).setValue(userName);
        custSheet.getRange(row, 3).setValue(realName);
        custSheet.getRange(row, 4).setValue(phone);
        // Col 6: 最後進站時間
        custSheet.getRange(row, 6).setValue(now);
        // Col 14: 顧客來源
        if (source) custSheet.getRange(row, 14).setValue(source);

        // 寫入更新資料安全 Log
        logCustomerActivity(
          userId,
          `${realName} (${userName})`,
          '✏️ 更新會員資料',
          `更新真實姓名: ${realName}, 電話: ${phone}, 來源: ${source || '未指定'}`,
          data.page || '個人資料/VIP守門牆',
          data.device || ''
        );
        break;
      }
    }

    if (!found && (!isRealLineId) && cleanPhone) {
      for (let i = 1; i < values.length; i++) {
        const rowPhone = String(values[i][3] || '').replace(/[-\s]/g, '').trim();
        const rowUid = String(values[i][0] || '').trim();
        if (rowPhone === cleanPhone && rowUid && !rowUid.startsWith('GUEST_') && rowUid !== 'TEST_USER_999') {
          userId = rowUid;
          found = true;
          const row = i + 1;
          if (userName) custSheet.getRange(row, 2).setValue(userName);
          custSheet.getRange(row, 3).setValue(realName);
          custSheet.getRange(row, 6).setValue(now);
          if (source) custSheet.getRange(row, 14).setValue(source);

          logCustomerActivity(
            userId,
            `${realName} (${userName})`,
            '✏️ 更新會員資料',
            `繼承帳號更新真實姓名: ${realName}, 電話: ${phone}`,
            data.page || '個人資料/VIP守門牆',
            data.device || ''
          );
          break;
        }
      }
    }

    if (!found) {
      custSheet.appendRow([
        userId,
        userName,
        realName,
        phone,
        now,                    // Col 5: 加入會員時間
        now,                    // Col 6: 最後進站時間
        data.page || 'VIP守門牆', // Col 7: 最後瀏覽頁面
        1,                      // Col 8: 進站累計天數
        1,                      // Col 9: 商品與頁面瀏覽次數
        0,                      // Col 10: 歷史訂單數
        0,                      // Col 11: 總消費金額
        '一般會員',              // Col 12: 會員等級
        '🟢 活躍',               // Col 13: 顧客生命狀態
        source || 'LINE 好友',   // Col 14: 顧客來源
        '',                     // Col 15: 常用超商取件門市
        '正常',                 // Col 16: 黑名單標記
        ''                      // Col 17: 賣家管理備註/標籤
      ]);

      // 寫入加入會員安全 Log
      logCustomerActivity(
        userId,
        `${realName} (${userName})`,
        '🎉 首次加入會員',
        `註冊真實姓名: ${realName}, 電話: ${phone}, 來源: ${source || 'LINE 好友'}`,
        data.page || 'VIP守門牆',
        data.device || ''
      );
    }

    // 同步會員本人真實姓名與電話回寫至該買家現有訂單之 W 欄 (col 23) 與 X 欄 (col 24)
    try {
      const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
      if (orderSheet) {
        ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);
        const orderValues = orderSheet.getDataRange().getValues();
        for (let o = 1; o < orderValues.length; o++) {
          if (String(orderValues[o][2] || '').trim() === userId) {
            orderSheet.getRange(o + 1, 23).setValue(realName);
            orderSheet.getRange(o + 1, 24).setValue(phone);
          }
        }
      }
    } catch (syncErr) {
      console.warn('同步訂單 W/X 欄警告:', syncErr);
    }

    return {
      success: true,
      message: '✅ 顧客資料已成功建立歸戶！',
      profile: {
        userId: userId,
        userName: userName,
        realName: realName,
        phone: rawPhone,
        source: source,
        joinedAt: now,
        lastVisitAt: now,
        memberTier: '一般會員',
        lifeStatus: '🟢 活躍'
      }
    };
  }

/**
  * 更新或建立顧客檔案歸戶
  * 欄位: ['LINE_User_ID', '最新暱稱', '真實姓名', '電話', '常用寄送地址', '歷史訂單數', '總消費金額', '首購日期', '最後下單日期', '黑名單標記', '顧客來源', '最近收件人姓名', '最近收件人電話']
  */
  function updateCustomerProfile(custSheet, orderData, totalAmount, now) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!custSheet) {
        custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
        if (!custSheet) {
          custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
          custSheet.appendRow([
            'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '加入會員時間', 
            '最後進站時間', '最後瀏覽商品/頁面', '進站累計天數', '商品與頁面瀏覽次數', '歷史訂單數', '總消費金額', 
            '會員等級', '顧客生命狀態', '顧客來源', '常用超商取件門市', '黑名單標記', '賣家管理備註/標籤'
          ]);
          custSheet.getRange(1, 1, 1, 17).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
          custSheet.getRange("D:D").setNumberFormat('@');
        }
      }

      ensureOrderAndCustomerSchema(ss, null, custSheet);

      const nowStr = now instanceof Date ? Utilities.formatDate(now, 'GMT+8', 'yyyy/MM/dd HH:mm:ss') : Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
      const userId = orderData.userId || 'LINE_GUEST';
      const userName = orderData.userName || '訪客';
      const realName = (orderData.realName || '').trim();
      const phone = formatPhoneAsText(orderData.phone);
      const source = (orderData.source || '').trim();
      const prodName = orderData.productName || '';
      const buyQty = Number(orderData.quantity) || 1;
      const data = custSheet.getDataRange().getValues();
      let found = false;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === userId) {
          found = true;
          const row = i + 1;
          if (userName) custSheet.getRange(row, 2).setValue(userName);
          if (realName && (!data[i][2] || String(data[i][2]).trim() === '')) {
            custSheet.getRange(row, 3).setValue(realName);
          }
          if (phone && (!data[i][3] || String(data[i][3]).trim() === '')) {
            custSheet.getRange(row, 4).setValue(phone);
          }
          // Col 6: 最後進站時間
          custSheet.getRange(row, 6).setValue(nowStr);
          // Col 7: 最後瀏覽商品/頁面
          if (prodName) {
            custSheet.getRange(row, 7).setValue(prodName);
          }
          
          const currentOrders = Number(data[i][9]) || 0;
          const currentSpend = Number(data[i][10]) || 0;
          const newOrders = currentOrders + 1;
          const newSpend = currentSpend + (Number(totalAmount) || 0);

          // Col 10: 歷史訂單數, Col 11: 總消費金額
          custSheet.getRange(row, 10).setValue(newOrders);
          custSheet.getRange(row, 11).setValue(newSpend);

          // Col 12: 會員等級自動試算 (尊重手動指定之「批發夥伴」；累積3000=VIP, 10000=VVIP)
          const currentTier = String(data[i][11] || '一般會員').trim();
          let autoTier = currentTier;
          if (currentTier !== '批發夥伴') {
            if (newSpend >= 10000) {
              autoTier = 'VVIP';
            } else if (newSpend >= 3000) {
              autoTier = (currentTier === 'VVIP' ? 'VVIP' : 'VIP');
            }
          }
          if (autoTier !== currentTier) {
            custSheet.getRange(row, 12).setValue(autoTier);
          }

          // Col 13: 顧客生命狀態 (活躍)
          custSheet.getRange(row, 13).setValue('🟢 活躍');

          // Col 14: 顧客來源
          if (source && (!data[i][13] || data[i][13] === '')) {
            custSheet.getRange(row, 14).setValue(source);
          }
          break;
        }
      }

      if (!found) {
        custSheet.appendRow([
          userId,
          userName,
          realName,
          phone,
          nowStr,       // 加入會員時間
          nowStr,       // 最後進站時間
          prodName,     // 最後瀏覽商品
          1,            // 進站累計天數
          1,            // 商品與頁面瀏覽次數
          1,            // 歷史訂單數
          Number(totalAmount) || 0, // 總消費金額
          '一般會員',    // 會員等級
          '🟢 活躍',     // 顧客生命狀態
          source,       // 顧客來源
          '',           // 常用超商取件門市
          '正常',       // 黑名單標記
          ''            // 賣家管理備註/標籤
        ]);
      }

      // 寫入送出訂單安全稽核 Log
      logCustomerActivity(
        userId,
        realName ? `${realName} (${userName})` : userName,
        '🛍️ 送出訂單',
        `下單商品: ${prodName || '未具名商品'}${buyQty > 1 ? ' x' + buyQty : ''}，金額 NT$ ${Number(totalAmount) || 0}`,
        'product.html',
        orderData.device || ''
      );
    } catch (err) {
      console.error('更新顧客歸戶失敗 (不影響訂單建立): ' + err.toString());
    }
  }

/**
  * 更新結帳時之最近收件人資訊與常用寄送地址
  * 嚴格只更新第 5 欄(E: 常用寄送地址)、第 12 欄(L: 最近收件人姓名)、第 13 欄(M: 最近收件人電話)
  * 絕不覆蓋第 3 欄 (C: 真實姓名) 或第 4 欄 (D: 電話)！
  */
  function updateCustomerRecipientInfo(custSheet, userId, recipientName, recipientPhone, deliveryInfo) {
    try {
      if (!custSheet) return;
      const cleanPhone = formatPhoneAsText(recipientPhone);
      const nowStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
      const data = custSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === userId) {
          const row = i + 1;
          // 第 15 欄 (O): 常用超商取件門市
          if (deliveryInfo) custSheet.getRange(row, 15).setValue(deliveryInfo);
          // 第 6 欄 (F): 最後進站時間
          custSheet.getRange(row, 6).setValue(nowStr);

          // 防呆：若該顧客 C欄(真實姓名) 或 D欄(電話) 原本完全空白，則順道補齊
          const currentRealName = String(data[i][2] || '').trim();
          const currentPhone = String(data[i][3] || '').trim();
          if (!currentRealName && recipientName) {
            custSheet.getRange(row, 3).setValue(recipientName);
          }
          if (!currentPhone && cleanPhone) {
            custSheet.getRange(row, 4).setValue(cleanPhone);
          }
          break;
        }
      }

      // 寫入完成結帳安全稽核 Log
      logCustomerActivity(
        userId,
        recipientName || '結帳買家',
        '💳 完成結帳',
        `收件人: ${recipientName || '未指定'} (${cleanPhone || ''})，取件門市/地址: ${deliveryInfo || ''}`,
        'my-orders.html',
        ''
      );
    } catch (err) {
      console.error('更新收件人資訊失敗: ' + err.toString());
    }
  }

/**
  * 相容舊函式命名：轉向呼叫 updateCustomerRecipientInfo
  */
  function updateCustomerAddress(custSheet, userId, name, phone) {
    updateCustomerRecipientInfo(custSheet, userId, name, phone, '');
  }

/**
  * 🔄 一鍵重新計算並同步【顧客歸戶】的所有統計數據
  * 根據【訂單明細】最新資料，自動為每位顧客精準重算：
  * - 歷史訂單數 (第 6 欄)
  * - 總消費金額 (第 7 欄)
  * - 首購日期 (第 8 欄)
  * - 最後下單日期 (第 9 欄)
  * 支援「LINE_User_ID」與「手機號碼」雙重智慧歸戶比對！
  */
  function recalculateCustomerStats() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    
    if (!orderSheet || !custSheet) {
      return { success: false, message: '找不到訂單或顧客工作表' };
    }

    // 自動執行欄位升級與資料同步 (將顧客表的真實姓名回寫至訂單表第 23 欄 W)
    ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);
    upgradeOrderAndCustomerSheets(ss);

    const orderData = orderSheet.getDataRange().getValues();
    const custData = custSheet.getDataRange().getValues();
    if (custData.length <= 1) {
      return { success: true, message: '顧客表無資料' };
    }

    // 1. 建立顧客映射表 (以 userId 為主 key，同時支援手機號碼比對)
    const statsMap = {};
    const phoneToUidMap = {};

    for (let c = 1; c < custData.length; c++) {
      const uid = String(custData[c][0] || '').trim();
      const phone = String(custData[c][3] || '').replace(/[-\s]/g, '').trim();
      if (uid) {
        statsMap[uid] = {
          orderCount: 0,
          totalSpend: 0,
          firstDate: null,
          lastDate: null
        };
        if (phone) phoneToUidMap[phone] = uid;
      }
    }

    // 2. 遍歷所有訂單進行統計累加
    for (let o = 1; o < orderData.length; o++) {
      const row = orderData[o];
      let uid = String(row[2] || '').trim(); // 第 3 欄 LINE_User_ID
      // 優先以第 24 欄 (X: 會員電話) 比對，其次以第 14 欄 (N: 收件人電話)
      const memberPhone = String(row[23] || '').replace(/[-\s]/g, '').trim();
      const recipientPhone = String(row[13] || '').replace(/[-\s]/g, '').trim();
      const phone = memberPhone || recipientPhone;
      const rawDate = row[1]; // 第 2 欄 下單時間
      const orderTime = rawDate instanceof Date ? rawDate : (rawDate ? new Date(rawDate) : null);
      const amount = Number(row[11]) || Number(row[9]) || 0; // 第 12 欄總金額或第 10 欄商品小計

      // 若 uid 未在顧客表直接對應，但手機號碼能對應，則智慧歸戶
      if (!statsMap[uid] && phone && phoneToUidMap[phone]) {
        uid = phoneToUidMap[phone];
      }

      if (statsMap[uid]) {
        statsMap[uid].orderCount += 1;
        statsMap[uid].totalSpend += amount;

        if (orderTime && !isNaN(orderTime.getTime())) {
          if (!statsMap[uid].firstDate || orderTime < statsMap[uid].firstDate) {
            statsMap[uid].firstDate = orderTime;
          }
          if (!statsMap[uid].lastDate || orderTime > statsMap[uid].lastDate) {
            statsMap[uid].lastDate = orderTime;
          }
        }
      }
    }

    // 3. 寫回【顧客歸戶】工作表 (17 欄規格)
    let updatedCount = 0;
    for (let c = 1; c < custData.length; c++) {
      const uid = String(custData[c][0] || '').trim();
      const stat = statsMap[uid];
      if (stat) {
        const row = c + 1;
        // Col 10: 歷史訂單數, Col 11: 總消費金額
        custSheet.getRange(row, 10).setValue(stat.orderCount);
        custSheet.getRange(row, 11).setValue(stat.totalSpend);

        // Col 12: 會員等級自動試算 (尊重批發夥伴；>=10000=VVIP, >=3000=VIP)
        const curTier = String(custData[c][11] || '一般會員').trim();
        let tier = curTier;
        if (curTier !== '批發夥伴') {
          if (stat.totalSpend >= 10000) tier = 'VVIP';
          else if (stat.totalSpend >= 3000) tier = (curTier === 'VVIP' ? 'VVIP' : 'VIP');
          else tier = '一般會員';
        }
        custSheet.getRange(row, 12).setValue(tier);

        // Col 5: 加入會員時間 (若原本空白且有首購紀錄則回補)
        if (stat.firstDate && !custData[c][4]) custSheet.getRange(row, 5).setValue(stat.firstDate);
        // Col 6: 最後進站時間 (若原本空白且有最後單紀錄則回補)
        if (stat.lastDate && !custData[c][5]) custSheet.getRange(row, 6).setValue(stat.lastDate);
        updatedCount++;
      }
    }

    logCustomerActivity('SYSTEM', '系統維護', '📊 重新統計會員數據', `已重新統計並同步 ${updatedCount} 位會員之歷史訂單與消費金額`, 'recalculateCustomerStats', 'GAS Server');

    Logger.log(`✅ 已完成 ${updatedCount} 位顧客之歷史訂單數據重新整理！`);
    return {
      success: true,
      message: `✅ 已重新統計並同步 ${updatedCount} 位顧客的歷史訂單與消費金額！`,
      updatedCount: updatedCount
    };
  }

/**
  * 📝 記錄會員安全與動態歷程 Log (安全稽核流水帳)
  * @param {string} userId - LINE User ID
  * @param {string} displayName - 姓名或暱稱
  * @param {string} eventType - 事件類型 (如 '🎉 首次加入會員', '👀 進站逛店', '🛍️ 送出訂單', '💳 完成結帳', '✏️ 資料變更')
  * @param {string} details - 詳細描述
  * @param {string} page - 頁面或管道
  * @param {string} device - 裝置與瀏覽器
  */
  function logCustomerActivity(userId, displayName, eventType, details, page, device) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
      const logHeaders = ['記錄時間', 'LINE_User_ID', '會員真實姓名/暱稱', '事件類型', '動態詳情', '來源頁面/管道', '裝置環境'];

      if (!logSheet) {
        logSheet = ss.insertSheet(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
        logSheet.appendRow(logHeaders);
        logSheet.getRange(1, 1, 1, logHeaders.length).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
        logSheet.getRange("A:A").setNumberFormat('yyyy/MM/dd HH:mm:ss');
        logSheet.getRange("B:B").setNumberFormat('@');
      }

      const nowStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
      logSheet.appendRow([
        nowStr,
        String(userId || 'GUEST'),
        String(displayName || '訪客'),
        String(eventType || '動態記錄'),
        String(details || ''),
        String(page || ''),
        String(device || '行動裝置/電腦')
      ]);
    } catch (err) {
      console.warn('寫入會員動態 Log 略過:', err);
    }
  }

/**
  * 查詢會員動態安全 Log (依時間倒序回傳最新記錄，供管理者後台查看)
  */
  function getCustomerLogs(limit = 100) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
      if (!logSheet) {
        return { success: true, logs: [], count: 0 };
      }
      const data = logSheet.getDataRange().getValues();
      if (data.length <= 1) {
        return { success: true, logs: [], count: 0 };
      }
      const logs = [];
      const maxRows = Math.min(Number(limit) || 100, data.length - 1);
      for (let r = data.length - 1; r >= data.length - maxRows; r--) {
        logs.push({
          timestamp: data[r][0] instanceof Date ? Utilities.formatDate(data[r][0], 'GMT+8', 'yyyy/MM/dd HH:mm:ss') : String(data[r][0]),
          userId: data[r][1],
          displayName: data[r][2],
          eventType: data[r][3],
          details: data[r][4],
          page: data[r][5],
          device: data[r][6]
        });
      }
      return { success: true, logs: logs, count: logs.length };
    } catch (err) {
      return { success: false, message: '讀取會員動態 Log 失敗: ' + err.toString() };
    }
  }

/**
  * 查詢待發送 1 對 1 結帳通知的顧客清單 (0 額度消耗推播後端)
  * 自動遍歷【訂單明細】與【顧客歸戶】，過濾出未結帳且採購確認的顧客名單
  */
  function getCustomerNotifyList() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    if (!orderSheet) {
      return { success: false, message: '訂單工作表不存在' };
    }

    // 確保 24 欄與 13 欄結構
    ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);

    const orderRows = orderSheet.getDataRange().getValues();
    const custRows = custSheet ? custSheet.getDataRange().getValues() : [];

    // 1. 建立顧客歸戶資料映射 (以 LINE_User_ID 為主 key，同時支援手機映射)
    const customerInfoMap = {};
    const phoneToCustomerMap = {};
    for (let c = 1; c < custRows.length; c++) {
      const uid = String(custRows[c][0] || '').trim();
      const phone = String(custRows[c][3] || '').replace(/[-\s]/g, '').trim();
      if (uid) {
        const info = {
          lineNick: String(custRows[c][1] || '').trim(),
          realName: String(custRows[c][2] || '').trim(),
          phone: String(custRows[c][3] || '').trim()
        };
        customerInfoMap[uid] = info;
        if (phone) phoneToCustomerMap[phone] = info;
      }
    }

    // 2. 歸組統計顧客未結帳訂單
    const customerGroups = {};

    for (let r = 1; r < orderRows.length; r++) {
      const row = orderRows[r];
      const userId = String(row[2] || '').trim();
      if (!userId) continue;

      const shipStatus = String(row[18] || '').trim();
      const checkoutId = String(row[21] || '').trim();
      const payStatus = String(row[16] || '').trim();
      const processNote = String(row[19] || '').trim();

      // 排除已結案、已完成取貨、已結帳或已付款完成者
      if (checkoutId) continue;
      if (payStatus === '已完成付款' || payStatus === '已付款待對帳') continue;
      if (shipStatus.includes('完成取貨') || shipStatus.includes('結案') || shipStatus.includes('已完成結帳') || shipStatus.includes('出貨')) {
        continue;
      }

      // 判斷商品狀態
      const isOut = shipStatus.includes('缺貨') || shipStatus.includes('斷貨') || shipStatus.includes('取消');
      const isSuccess = !isOut && (shipStatus.includes('成功') || shipStatus.includes('採買') || shipStatus.includes('通知'));
      const isRegistering = !isOut && !isSuccess;

      const qty = Number(row[7]) || 1;
      const unitPrice = Number(row[8]) || 0;
      const subtotal = Number(row[9]) || (unitPrice * qty);
      // 優先以會員登記電話 (X: col 24)，其次以收件人電話 (N: col 14)
      const phone = String(row[23] || row[13] || '').replace(/[-\s]/g, '').trim();

      if (!customerGroups[userId]) {
        const profile = customerInfoMap[userId] || (phone && phoneToCustomerMap[phone]) || {};
        // 姓名優先使用會員真實姓名 (row[22] W 欄或 profile.realName)，絕不以包裹收件人姓名 (row[12]) 混充會員姓名！
        const memberDisplayName = profile.realName || String(row[22] || '').trim() || profile.lineNick || String(row[3] || '').trim() || '顧客';
        const memberContactPhone = profile.phone || String(row[23] || '').trim() || String(row[13] || '').trim();

        customerGroups[userId] = {
          userId: userId,
          name: memberDisplayName,
          phone: memberContactPhone,
          successCount: 0,
          outOfStockCount: 0,
          pendingCount: 0,
          totalAmount: 0,
          items: [],
          lastNotifiedTime: null,
          isNotified: false
        };
      }

      const group = customerGroups[userId];

      // 檢查處理備註是否有通知時間標記
      if (processNote.includes('已發送1對1結帳通知')) {
        group.isNotified = true;
        const match = processNote.match(/已發送1對1結帳通知\s*\(([^)]+)\)/);
        if (match) {
          group.lastNotifiedTime = match[1];
        }
      }

      group.items.push({
        orderId: String(row[0] || '').trim(),
        productId: String(row[4] || '').trim(),
        productName: String(row[5] || '').trim(),
        spec: String(row[6] || '').trim(),
        quantity: qty,
        unitPrice: unitPrice,
        subtotal: subtotal,
        status: isSuccess ? 'success' : (isOut ? 'out' : 'pending'),
        shippingStatus: shipStatus
      });

      if (isSuccess) {
        group.successCount += qty;
        group.totalAmount += subtotal;
      } else if (isOut) {
        group.outOfStockCount += qty;
      } else if (isRegistering) {
        group.pendingCount += qty;
      }
    }

    // 3. 過濾與分類顧客
    const list = [];
    Object.values(customerGroups).forEach(group => {
      if (group.successCount === 0 && group.outOfStockCount > 0 && group.pendingCount === 0) {
        group.type = 'all_out'; // 全數缺貨斷貨
        list.push(group);
      } else if (group.successCount > 0) {
        group.type = 'checkout'; // 有採購成功品項，可結帳
        list.push(group);
      } else if (group.pendingCount > 0) {
        group.type = 'pending'; // 尚有連線登記未確認商品
        list.push(group);
      }
    });

    // 4. 排序規則：採購成功總金額由大到小排序 (高客單價大單買家優先通知)
    list.sort((a, b) => {
      if (a.type === 'checkout' && b.type === 'checkout') {
        return b.totalAmount - a.totalAmount;
      }
      if (a.type === 'checkout') return -1;
      if (b.type === 'checkout') return 1;
      return 0;
    });

    // 5. 彙整統計數據
    const checkoutList = list.filter(c => c.type === 'checkout');
    const stats = {
      totalCheckoutCustomers: checkoutList.length,
      totalSuccessItems: checkoutList.reduce((sum, c) => sum + c.successCount, 0),
      totalPendingAmount: checkoutList.reduce((sum, c) => sum + c.totalAmount, 0)
    };

    return {
      success: true,
      stats: stats,
      customers: list
    };
  }

/**
  * 標記顧客名下所有當期未結帳訂單為「通知結帳」，並記錄通知發送時間
  */
  function handleMarkCustomerNotified(postData) {
    const data = (postData && postData.data) || postData || {};
    const userId = data.userId;
    if (!userId) return { success: false, message: '缺少 userId' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    if (!orderSheet) return { success: false, message: '訂單工作表不存在' };

    const rows = orderSheet.getDataRange().getValues();
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd HH:mm');
    let updatedCount = 0;

    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][2] || '').trim() === userId) {
        const shipStatus = String(rows[r][18] || '').trim();
        const checkoutId = String(rows[r][21] || '').trim();
        if (!checkoutId && !shipStatus.includes('完成取貨') && !shipStatus.includes('結案')) {
          const rowIndex = r + 1;
          // 若狀態為採購成功，狀態更新為「通知結帳」
          if (shipStatus === '採購成功' || shipStatus === '連線登記') {
            orderSheet.getRange(rowIndex, 19).setValue('通知結帳');
          }
          // 追記備註
          const oldNote = String(rows[r][19] || '').trim();
          const notifyTag = `已發送1對1結帳通知 (${nowStr})`;
          let newNote = oldNote;
          if (oldNote.includes('已發送1對1結帳通知')) {
            newNote = oldNote.replace(/已發送1對1結帳通知\s*\([^)]*\)/, notifyTag);
          } else {
            newNote = oldNote ? `${oldNote} ｜ ${notifyTag}` : notifyTag;
          }
          orderSheet.getRange(rowIndex, 20).setValue(newNote);
          updatedCount++;
        }
      }
    }

    return {
      success: true,
      message: `已成功更新 ${updatedCount} 筆訂單備註為已通知`,
      notifiedTime: nowStr
    };
  }

/**
  * 依據 LINE User ID 反查顧客歸戶檔案 (取得 LINE 暱稱、真實姓名、手機)
  */
  function getCustomerProfileById(userId) {
    if (!userId || String(userId).startsWith('GUEST_') || userId === 'LINE_GUEST' || userId === 'TEST_USER_999') {
      return null;
    }
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS || '顧客歸戶');
      if (!custSheet) return null;
      const custData = custSheet.getDataRange().getValues();
      for (let c = 1; c < custData.length; c++) {
        if (String(custData[c][0] || '').trim() === String(userId).trim()) {
          return {
            lineName: String(custData[c][1] || '').trim(),
            realName: String(custData[c][2] || '').trim(),
            phone: String(custData[c][3] || '').trim()
          };
        }
      }
    } catch (err) {
      console.warn('getCustomerProfileById 查詢失敗:', err);
    }
    return null;
  }

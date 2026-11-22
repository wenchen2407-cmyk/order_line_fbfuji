/**
 * ==============================================================================
 * 【03_OrderService.gs】訂單管理與出貨服務模組 (Order & Procurement)
 * ==============================================================================
 * 說明：專門處理買家下單保留、加單合併、結帳出貨、匯款回報、採購作業站與包裹狀態。
 * ==============================================================================
 */

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

      // 2. 檢查收單截止時間防呆 (優先以各商品指定結單時間為準，未填則採用系統設定 ORDER_DEADLINE)
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

      const targetSpec = String(orderData.spec || '單一規格').trim();
      let specsList = [];
      let specsContainer = null;
      try {
        const parsed = JSON.parse(targetProduct[6]);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed) && parsed.variants) {
          specsContainer = parsed;
          specsList = parsed.variants;
        } else if (Array.isArray(parsed)) {
          specsList = parsed;
        }
      } catch(e) {
        specsList = targetProduct[6] ? [targetProduct[6]] : [];
      }

      let targetSpecObj = null;
      let targetSpecIndex = -1;
      let unitPrice = Number(targetProduct[4]); // 預設為主商品連線代購價

      if (Array.isArray(specsList)) {
        for (let s = 0; s < specsList.length; s++) {
          const item = specsList[s];
          const sName = (item && typeof item === 'object') ? String(item.name || '').trim() : String(item || '').trim();
          if (sName === targetSpec) {
            targetSpecObj = item;
            targetSpecIndex = s;
            if (item && typeof item === 'object') {
              if (item.price !== undefined && !isNaN(Number(item.price)) && Number(item.price) > 0) {
                unitPrice = Number(item.price);
              }
            }
            break;
          }
        }
      }

      const buyQty = Number(orderData.quantity) || 1;

      // 檢查規格庫存與總庫存
      if (targetSpecObj && typeof targetSpecObj === 'object') {
        const sStock = Number(targetSpecObj.stock);
        const isUnlimited = Boolean(targetSpecObj.isUnlimited || sStock >= 9999);
        if (!isUnlimited && sStock < buyQty) {
          return { success: false, message: `⚠️ 款式「${targetSpec}」庫存不足！目前剩餘數量：${sStock}` };
        }
      } else {
        if (currentStock < buyQty) {
          return { success: false, message: '庫存不足！目前剩餘數量：' + currentStock };
        }
      }

      // 扣減指定規格庫存與總庫存
      if (targetSpecObj && typeof targetSpecObj === 'object' && targetSpecIndex >= 0) {
        const sStock = Number(targetSpecObj.stock);
        const isUnlimited = Boolean(targetSpecObj.isUnlimited || sStock >= 9999);
        if (!isUnlimited) {
          targetSpecObj.stock = Math.max(0, sStock - buyQty);
          specsList[targetSpecIndex] = targetSpecObj;
          if (specsContainer) {
            specsContainer.variants = specsList;
            prodSheet.getRange(productRowIndex, 7).setValue(JSON.stringify(specsContainer));
          } else {
            prodSheet.getRange(productRowIndex, 7).setValue(JSON.stringify(specsList));
          }
        }
      }

      const newStock = Math.max(0, currentStock - buyQty);
      prodSheet.getRange(productRowIndex, 6).setValue(newStock);
      if (newStock <= 0) {
        prodSheet.getRange(productRowIndex, 10).setValue('已售完');
      }

      // 3. 檢查同買家是否已登記同商品之同規格（尚未結帳），若有則自動累加數量（合單）
      let targetUserId = String(orderData.userId || '').trim();
      const cleanPhone = String(orderData.phone || '').replace(/[-\s]/g, '').trim();

      // 🛡️ 徹底消滅 GUEST：若傳入 GUEST 或空值，自動依手機號碼查詢是否已有真實 LINE ID
      if (!targetUserId || targetUserId.startsWith('GUEST_') || targetUserId === 'TEST_USER_999' || targetUserId === 'LINE_GUEST') {
        if (cleanPhone && custSheet) {
          const custRows = custSheet.getDataRange().getValues();
          for (let c = 1; c < custRows.length; c++) {
            const rowPhone = String(custRows[c][3] || '').replace(/[-\s]/g, '').trim();
            const rowUid = String(custRows[c][0] || '').trim();
            if (rowPhone === cleanPhone && rowUid && !rowUid.startsWith('GUEST_') && rowUid !== 'TEST_USER_999') {
              targetUserId = rowUid;
              orderData.userId = rowUid;
              break;
            }
          }
        }
      }

      const orderDataRange = orderSheet.getDataRange();
      const orderValues = orderDataRange.getValues();
      
      let existingOrderRow = -1;
      let existingOrderId = '';
      let existingQty = 0;
      let existingNote = '';
      let existingProcessNote = '';

      // 從最新（最後一列）往回查找，確保匹配當期最近未結帳項目
      if (targetUserId && !targetUserId.startsWith('GUEST_') && targetUserId !== 'LINE_GUEST' && targetUserId !== 'TEST_USER_999') {
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

        // 補全可能遺漏的 W欄(23:真實姓名) 與 X欄(24:會員電話)
        try {
          const currentRealNameInOrder = String(orderSheet.getRange(existingOrderRow, 23).getValue() || '').trim();
          const currentMemberPhoneInOrder = String(orderSheet.getRange(existingOrderRow, 24).getValue() || '').trim();
          const targetRealName = (orderData.realName || '').trim();
          const targetMemberPhone = formatPhoneAsText(orderData.phone);
          if (!currentRealNameInOrder && targetRealName) {
            orderSheet.getRange(existingOrderRow, 23).setValue(targetRealName);
          }
          if (!currentMemberPhoneInOrder && targetMemberPhone) {
            orderSheet.getRange(existingOrderRow, 24).setValue(targetMemberPhone);
          }
        } catch (colSyncErr) {
          console.warn('合單補全真實姓名電話警告:', colSyncErr);
        }

        // 若有新備註則追加
        if (orderData.note && orderData.note.trim()) {
          const mergedNote = existingNote ? `${existingNote}；${orderData.note.trim()}` : orderData.note.trim();
          orderSheet.getRange(existingOrderRow, 16).setValue(mergedNote);
        }

        // 判斷商品供應模式 (現貨即時採購成功，無須等待現場連線登記)
        const prodCategory = String(targetProduct[2] || '').trim();
        const isSpotGoods = prodCategory.includes('現貨');

        // 在處理備註(T欄第20欄)自動標記加單紀錄，方便賣家清楚追蹤
        const spotTag = isSpotGoods ? '[⚡現貨]' : '';
        const mergeLog = `[加單+${buyQty}${spotTag} (${timeFormatted})]`;
        const newProcessNote = existingProcessNote ? `${existingProcessNote} ${mergeLog}` : mergeLog;
        orderSheet.getRange(existingOrderRow, 20).setValue(newProcessNote);

        // 若商品為現貨，確保訂單狀態為「採購成功」
        if (isSpotGoods) {
          const currentShipStatus = String(orderSheet.getRange(existingOrderRow, 19).getValue() || '').trim();
          if (currentShipStatus === '連線登記' || !currentShipStatus) {
            orderSheet.getRange(existingOrderRow, 19).setValue('採購成功');
          }
        }

      } else {
        // 🌟 新成立訂單
        const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss');
        const randomSuffix = Math.floor(100 + Math.random() * 900);
        finalOrderId = 'OD' + timeStr + randomSuffix;
        finalQty = buyQty;
        finalSubtotal = unitPrice * buyQty;
        const shippingFee = 0; // 連線期間運費先設為 0，出貨結帳時統一合併計算
        const totalAmount = finalSubtotal;

        // 判斷商品供應模式 (現貨即時採購成功，無須現場連線登記)
        const prodCategory = String(targetProduct[2] || '').trim();
        const isSpotGoods = prodCategory.includes('現貨');
        const initialShipStatus = isSpotGoods ? '採購成功' : '連線登記';
        const initialProcessNote = isSpotGoods ? '[⚡ 現貨即時鎖定]' : '';

        // 確保訂單表具備 24 欄結構 (M/N 欄保留給出貨結帳收件人，W/X 欄記錄會員真實身分)
        ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);

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
          '',                                        // 第 13 欄 (M: 收件人姓名 - 待回國結帳時填寫包裹收件人)
          '',                                        // 第 14 欄 (N: 收件人電話 - 待回國結帳時填寫包裹收件人電話)
          '[待出貨結帳填寫]',                         // 第 15 欄 (O: 取件方式與地址)
          orderData.note || '',                      // 第 16 欄 (P: 訂單備註)
          '未結帳',                                   // 第 17 欄 (Q: 付款狀態)
          '',                                        // 第 18 欄 (R: 匯款後五碼)
          initialShipStatus,                         // 第 19 欄 (S: 訂單採購/出貨處理狀態: 現貨直接採購成功)
          initialProcessNote,                        // 第 20 欄 (T: 處理備註)
          '',                                        // 第 21 欄 (U: 包裹追蹤編號)
          '',                                        // 第 22 欄 (V: 結帳編號)
          orderData.realName || '',                  // 第 23 欄 (W: 真實姓名 - LINE帳號會員本人登記真實姓名)
          formatPhoneAsText(orderData.phone)         // 第 24 欄 (X: 會員電話 - LINE帳號會員本人登記電話)
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

      // 🔔 即時發送【買家下單登記 / 加單】Telegram 通知 (稱謂統稱「訪客稱呼」)
      try {
        const rawSettings = getRawSystemSettingsInternal();
        if (rawSettings.TELEGRAM_NOTIFY_ORDERS !== 'NO') {
          let vLineName = orderData.userName || '';
          let vRealName = orderData.realName || '';
          let vPhone = orderData.phone || '';
          if (!vRealName || !vPhone) {
            const profile = getCustomerProfileById(targetUserId || orderData.userId);
            if (profile) {
              if (!vLineName) vLineName = profile.lineName;
              if (!vRealName) vRealName = profile.realName;
              if (!vPhone) vPhone = profile.phone;
            }
          }
          const visitorDisplay = formatVisitorDisplayName(vLineName, vRealName, vPhone);
          const timeFormatted = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
          const orderMsg = 
            `🛍️ <b>【買家下單登記】</b>\n` +
            `• <b>訪客稱呼</b>：${escapeHtmlText(visitorDisplay)}\n` +
            `• <b>登記商品</b>：${escapeHtmlText(targetProduct[1])}\n` +
            `• <b>選購規格</b>：${escapeHtmlText(orderData.spec || '單一規格')}\n` +
            `• <b>登記數量</b>：${finalQty} 件 (${isMerged ? '加單合併累加' : '新登記保留'})\n` +
            `• <b>商品金額</b>：NT$ ${finalSubtotal.toLocaleString()}\n` +
            `• <b>訂單編號</b>：<code>${finalOrderId}</code>\n` +
            `• <b>登記時間</b>：${timeFormatted}`;
          sendTelegramMessage(orderMsg);
        }
      } catch (tgErr) {
        console.warn('發送新訂單 Telegram 通知略過:', tgErr);
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

    // 確保 24 欄與 13 欄擴充欄位存在
    ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);

    // 反查會員本人登記之真實姓名與電話，確保寫入 W 欄與 X 欄
    let memberRealName = '';
    let memberPhone = '';
    if (custSheet) {
      const custValues = custSheet.getDataRange().getValues();
      for (let c = 1; c < custValues.length; c++) {
        if (String(custValues[c][0] || '').trim() === userId) {
          memberRealName = String(custValues[c][2] || '').trim();
          memberPhone = String(custValues[c][3] || '').trim();
          break;
        }
      }
    }

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

      // 同步/維持會員本人真實姓名 (第 23 欄 W) 與會員電話 (第 24 欄 X)
      try {
        const existingW = String(orderSheet.getRange(r, 23).getValue() || '').trim();
        const existingX = String(orderSheet.getRange(r, 24).getValue() || '').trim();
        if (!existingW && memberRealName) {
          orderSheet.getRange(r, 23).setValue(memberRealName);
        }
        if (!existingX && memberPhone) {
          orderSheet.getRange(r, 24).setValue(formatPhoneAsText(memberPhone));
        }
      } catch (e) {}
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

      // 同步/維持會員本人真實姓名 (第 23 欄 W) 與會員電話 (第 24 欄 X)
      try {
        const existingW = String(orderSheet.getRange(r, 23).getValue() || '').trim();
        const existingX = String(orderSheet.getRange(r, 24).getValue() || '').trim();
        if (!existingW && memberRealName) {
          orderSheet.getRange(r, 23).setValue(memberRealName);
        }
        if (!existingX && memberPhone) {
          orderSheet.getRange(r, 24).setValue(formatPhoneAsText(memberPhone));
        }
      } catch (e) {}
    }

    // 更新顧客歸戶 (更新最近包裹收件人姓名/電話與常用地址，嚴格不覆蓋會員真實姓名與會員電話)
    if (custSheet) {
      updateCustomerRecipientInfo(custSheet, userId, checkoutData.recipientName, checkoutData.recipientPhone, deliveryInfo);
    }

    // 🔔 即時發送【買家完成出貨結帳】Telegram 通知 (稱謂統稱「訪客稱呼」)
    try {
      const rawSettings = getRawSystemSettingsInternal();
      if (rawSettings.TELEGRAM_NOTIFY_PAYMENTS !== 'NO') {
        let vLineName = checkoutData.userName || '';
        let vRealName = checkoutData.recipientName || memberRealName || '';
        let vPhone = checkoutData.recipientPhone || memberPhone || '';
        if (!vRealName || !vPhone) {
          const profile = getCustomerProfileById(userId);
          if (profile) {
            if (!vLineName) vLineName = profile.lineName;
            if (!vRealName) vRealName = profile.realName;
            if (!vPhone) vPhone = profile.phone;
          }
        }
        const visitorDisplay = formatVisitorDisplayName(vLineName, vRealName, vPhone);
        const timeFormatted = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
        const checkoutMsg = 
          `🛒 <b>【買家完成出貨結帳】</b>\n` +
          `• <b>訪客稱呼</b>：${escapeHtmlText(visitorDisplay)}\n` +
          `• <b>結帳編號</b>：<code>${checkoutId}</code>\n` +
          `• <b>配送方式</b>：${escapeHtmlText(checkoutData.shippingMethodName || method)}\n` +
          `• <b>採購商品</b>：共 ${successfulIndices.length} 件\n` +
          `• <b>結帳總額</b>：NT$ ${finalTotalAmount.toLocaleString()} (含運 NT$ ${shippingFee})\n` +
          `• <b>收件姓名</b>：${escapeHtmlText(checkoutData.recipientName || '')}\n` +
          `• <b>收件電話</b>：${escapeHtmlText(checkoutData.recipientPhone || '')}\n` +
          `• <b>收件門市/地址</b>：${escapeHtmlText(deliveryInfo)}\n` +
          `• <b>結帳時間</b>：${timeFormatted}`;
        sendTelegramMessage(checkoutMsg);
      }
    } catch (tgErr) {
      console.warn('發送出貨結帳 Telegram 通知略過:', tgErr);
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
      try {
        logCustomerActivity(
          data.userId || '',
          data.userName || '會員買家',
          '💰 回報匯款',
          `方式: ${payMethod}，轉帳時間: ${transferTime}${lastFive ? '，後五碼: ' + lastFive : ''}${data.note ? '，備註: ' + data.note : ''}`,
          'my-orders.html',
          data.device || ''
        );
      } catch (e) {
        console.warn('記錄回報匯款 Log 失敗:', e);
      }

      // 🔔 即時發送【買家匯款回報】Telegram 通知 (稱謂統稱「訪客稱呼」)
      try {
        const rawSettings = getRawSystemSettingsInternal();
        if (rawSettings.TELEGRAM_NOTIFY_PAYMENTS !== 'NO') {
          let vLineName = data.userName || '';
          let vRealName = data.realName || '';
          let vPhone = data.phone || '';
          if (!vRealName || !vPhone) {
            const profile = getCustomerProfileById(data.userId);
            if (profile) {
              if (!vLineName) vLineName = profile.lineName;
              if (!vRealName) vRealName = profile.realName;
              if (!vPhone) vPhone = profile.phone;
            }
          }
          const visitorDisplay = formatVisitorDisplayName(vLineName, vRealName, vPhone);
          const nowFormatted = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
          const payMsg = 
            `💳 <b>【買家匯款回報】</b>\n` +
            `• <b>訪客稱呼</b>：${escapeHtmlText(visitorDisplay)}\n` +
            `• <b>付款方式</b>：${escapeHtmlText(payMethod)}\n` +
            `• <b>匯款後五碼</b>：${lastFive ? '<code>' + escapeHtmlText(String(data.lastFiveDigits).trim()) + '</code>' : '(未填寫)'}\n` +
            `• <b>轉帳時間</b>：${escapeHtmlText(transferTime)}\n` +
            (data.note ? `• <b>買家備註</b>：${escapeHtmlText(data.note)}\n` : '') +
            `• <b>回報時間</b>：${nowFormatted}`;
          sendTelegramMessage(payMsg);
        }
      } catch (tgErr) {
        console.warn('發送匯款回報 Telegram 通知略過:', tgErr);
      }

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
      try {
        logCustomerActivity(
          userId,
          data.userName || '會員買家',
          '📦 確認取貨',
          `買家確認已完成取貨 (${updatedCount} 筆商品)，訂單正式歸檔至歷史紀錄`,
          'my-orders.html',
          data.device || ''
        );
      } catch (e) {
        console.warn('記錄確認取貨 Log 失敗:', e);
      }
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

    // 確保 24 欄與 13 欄結構
    ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);

    // 取得顧客檔案 (含最近收件人姓名/電話)
    let customerProfile = { 
      userId: userId, 
      realName: '', 
      phone: '', 
      defaultAddress: '',
      lastRecipientName: '',
      lastRecipientPhone: ''
    };
    let memberRealName = '';
    let memberPhone = '';

    if (custSheet) {
      const custData = custSheet.getDataRange().getValues();
      for (let c = 1; c < custData.length; c++) {
        if (String(custData[c][0] || '').trim() === userId) {
          customerProfile.userName = custData[c][1];
          customerProfile.realName = custData[c][2];
          customerProfile.phone = custData[c][3];
          customerProfile.joinedAt = custData[c][4] || '';
          customerProfile.lastVisitAt = custData[c][5] || '';
          customerProfile.memberTier = custData[c][11] || '一般會員';
          customerProfile.defaultAddress = custData[c][14] || '';
          customerProfile.lastRecipientName = custData[c][2] || '';
          customerProfile.lastRecipientPhone = custData[c][3] || '';
          memberRealName = String(custData[c][2] || '').trim();
          memberPhone = String(custData[c][3] || '').trim();
          break;
        }
      }
    }

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
      if (String(rows[i][2] || '').trim() === userId) {
        const sStatus = String(rows[i][18] || '').trim();
        const cid = String(rows[i][21] || '').trim();
        if (cid && (sStatus.includes('已完成取貨') || sStatus.includes('結案') || sStatus.includes('已取消'))) {
          closedCheckoutIds.add(cid);
        }
      }
    }

    // 建立商品供應模式速查字典
    const prodModeMap = {};
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    if (prodSheet) {
      const pData = prodSheet.getDataRange().getValues();
      for (let p = 1; p < pData.length; p++) {
        const pid = String(pData[p][0] || '').trim();
        const cat = String(pData[p][2] || '').trim();
        if (pid) {
          let mode = '連線';
          if (cat.includes('現貨')) mode = '現貨';
          else if (cat.includes('預購')) mode = '預購';
          else mode = '連線';
          prodModeMap[pid] = { category: cat || '✈️ 連線', supplyMode: mode };
        }
      }
    }

    let needOrderSheetSync = false;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[2] || '').trim() === userId) {
        const pInfo = prodModeMap[row[4]] || { category: '✈️ 連線', supplyMode: '連線' };
        const orderRealName = String(row[22] || '').trim();
        const orderMemberPhone = String(row[23] || '').trim();

        // 若【訂單明細】W/X 欄空白，但【顧客歸戶】已有資料，自動同步回寫！
        if ((!orderRealName && memberRealName) || (!orderMemberPhone && memberPhone)) {
          try {
            if (!orderRealName && memberRealName) {
              orderSheet.getRange(i + 1, 23).setValue(memberRealName);
            }
            if (!orderMemberPhone && memberPhone) {
              orderSheet.getRange(i + 1, 24).setValue(formatPhoneAsText(memberPhone));
            }
            needOrderSheetSync = true;
          } catch (e) {}
        }

        const effectiveRealName = orderRealName || memberRealName || '';
        const effectiveMemberPhone = orderMemberPhone || memberPhone || '';
        const recipientName = String(row[12] || '').trim();
        const recipientPhone = String(row[13] || '').trim();

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
          recipientName: recipientName,
          recipientPhone: recipientPhone,
          phone: recipientPhone || effectiveMemberPhone,
          realName: effectiveRealName,
          memberPhone: effectiveMemberPhone,
          deliveryAddress: formatDeliveryText(row[14]),
          note: row[15],
          paymentStatus: row[16],
          lastFive: row[17],
          shippingStatus: row[18],
          processNote: row[19] || '',
          trackingNumber: String(row[20] || '').trim(), // 第 21 欄 U 欄：包裹追蹤編號
          checkoutId: String(row[21] || '').trim(),       // 第 22 欄 V 欄：結帳編號 (Checkout_ID)
          category: pInfo.category,                      // 供應模式/分類
          supplyMode: pInfo.supplyMode                   // 標準供應模式 (現貨 / 預購 / 連線)
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

/**
  * 查詢已完成結帳出貨對帳包裹清單 (管理中心第 4 站真實後端)
  * 遍歷【訂單明細】，依結帳編號 (checkoutId, 第 22 欄) 分組為已結帳待出貨包裹
  */
  function getCheckoutParcels() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    if (!orderSheet) {
      return { success: false, message: '訂單工作表不存在' };
    }

    ensureOrderAndCustomerSchema(ss, orderSheet, custSheet);

    // 預載商品清單圖片對照 Map (精確對齊第 8 欄 H 欄: 封面主圖網址)
    const prodMap = {};
    if (prodSheet) {
      const prodValues = prodSheet.getDataRange().getValues();
      for (let p = 1; p < prodValues.length; p++) {
        const pId = String(prodValues[p][0] || '').trim();
        if (pId) {
          let img = String(prodValues[p][7] || '').trim(); // H 欄 (8): 封面主圖網址
          if (!img || img.startsWith('[') || img.startsWith('{')) {
            try {
              const parsed = JSON.parse(img);
              if (Array.isArray(parsed) && parsed.length > 0) {
                img = typeof parsed[0] === 'string' ? parsed[0] : (parsed[0].url || parsed[0].imageUrl || '');
              }
            } catch (e) { img = ''; }
          }
          if (!img) {
            const allImgsRaw = String(prodValues[p][13] || '').trim(); // N 欄 (14): 所有圖片清單(JSON)
            if (allImgsRaw.startsWith('[')) {
              try {
                const parsedAll = JSON.parse(allImgsRaw);
                if (Array.isArray(parsedAll) && parsedAll.length > 0) {
                  img = typeof parsedAll[0] === 'string' ? parsedAll[0] : (parsedAll[0].url || '');
                }
              } catch (e) {}
            }
          }
          prodMap[pId] = {
            imageUrl: img
          };
        }
      }
    }

    const orderRows = orderSheet.getDataRange().getValues();
    const parcelMap = {};

    for (let r = 1; r < orderRows.length; r++) {
      const row = orderRows[r];
      const checkoutId = String(row[21] || '').trim(); // 第 22 欄 (V 欄)：結帳編號
      const shipStatus = String(row[18] || '').trim(); // 第 19 欄 (S 欄)：採購出貨狀態

      // 必須有 checkoutId
      if (!checkoutId) continue;
      const isCompleted = shipStatus.includes('完成取貨') || shipStatus.includes('結案');

      const userId = String(row[2] || '').trim();
      const qty = Number(row[7]) || 1;
      const unitPrice = Number(row[8]) || 0;
      const subtotal = Number(row[9]) || (unitPrice * qty);
      const shippingFee = Number(row[10]) || 0;
      const totalWithFee = Number(row[11]) || (subtotal + shippingFee);

      if (!parcelMap[checkoutId]) {
        const recipientName = String(row[12] || '').trim() || String(row[22] || '').trim() || String(row[3] || '').trim() || '顧客';
        const recipientPhone = String(row[13] || '').trim() || String(row[23] || '').trim();
        const deliveryInfo = String(row[14] || '').trim(); // 取件方式與地址 (O 欄第 15 欄)

        // 精確解析配送方式與門市店名 (避免取貨方式被誤當成超商門市)
        let deliveryMethod = '超商取貨';
        let storeBranch = '';
        if (deliveryInfo.includes(']')) {
          const parts = deliveryInfo.split(']');
          deliveryMethod = parts[0].replace('[', '').trim();
          storeBranch = parts.slice(1).join(']').trim();
        } else {
          storeBranch = deliveryInfo;
        }

        let fullStoreName = storeBranch || '超商門市';
        if (deliveryMethod.includes('7-11')) {
          if (!fullStoreName.includes('7-11') && !fullStoreName.includes('7-ELEVEN')) {
            fullStoreName = `7-11 ${storeBranch}${storeBranch.endsWith('門市') || storeBranch.endsWith('店') ? '' : '門市'}`;
          }
        } else if (deliveryMethod.includes('全家')) {
          if (!fullStoreName.includes('全家')) {
            fullStoreName = `全家 ${storeBranch}${storeBranch.endsWith('門市') || storeBranch.endsWith('店') ? '' : '店'}`;
          }
        } else if (deliveryMethod.includes('黑貓') || deliveryMethod.includes('宅配') || deliveryMethod.includes('郵局')) {
          fullStoreName = deliveryMethod;
        }

        let fullStoreAddr = '';
        if (deliveryMethod.includes('黑貓') || deliveryMethod.includes('宅配') || deliveryMethod.includes('郵局')) {
          fullStoreAddr = storeBranch ? `${deliveryMethod} ｜ ${storeBranch}` : deliveryMethod;
        } else {
          fullStoreAddr = storeBranch ? `${deliveryMethod} ｜ ${fullStoreName}` : deliveryMethod;
        }

        const rawPayStatus = String(row[16] || '').trim();
        const isCOD = deliveryInfo.includes('取貨付款') || deliveryInfo.includes('貨到付款') || rawPayStatus.includes('貨到付款') || rawPayStatus.includes('取貨付款');

        parcelMap[checkoutId] = {
          coCode: checkoutId,
          userId: userId,
          name: recipientName,
          phone: recipientPhone,
          lineName: String(row[3] || '').trim(),
          realName: String(row[22] || '').trim(),
          memberPhone: String(row[23] || '').trim(),
          deliveryMethod: deliveryMethod,
          storeBranch: storeBranch,
          store: fullStoreName,
          storeAddr: fullStoreAddr,
          deliveryInfo: deliveryInfo,
          isCOD: isCOD,
          payStatus: rawPayStatus || (isCOD ? '貨到付款待出貨' : '待付款'),
          last5: String(row[17] || '').replace(/^'/, '').trim(),
          shippingStatus: shipStatus || '已完成結帳',
          isCompleted: isCompleted,
          note: String(row[15] || '').trim(),
          processNote: String(row[19] || '').trim(),
          trackingCode: String(row[20] || '').trim(), // 第 21 欄 (U 欄)：包裹追蹤單號
          shippingFee: shippingFee,
          total: totalWithFee,
          items: []
        };
      } else {
        if (shippingFee > 0 && parcelMap[checkoutId].shippingFee === 0) {
          parcelMap[checkoutId].shippingFee = shippingFee;
        }
        parcelMap[checkoutId].total += subtotal;
      }

      const prodInfo = prodMap[String(row[4] || '').trim()] || {};
      parcelMap[checkoutId].items.push({
        orderId: String(row[0] || '').trim(),
        productId: String(row[4] || '').trim(),
        productName: String(row[5] || '').trim(),
        spec: String(row[6] || '').trim(),
        quantity: qty,
        unitPrice: unitPrice,
        subtotal: subtotal,
        imageUrl: prodInfo.imageUrl || '',
        status: 'success'
      });
    }

    const parcels = Object.values(parcelMap);
    return {
      success: true,
      count: parcels.length,
      parcels: parcels
    };
  }

  /**
  * 查詢現場採購與顧客訂單配貨清單 (管理中心第 1 站與第 2 站獨立資料來源 API)
  * 100% 依 Google Sheets【訂單明細】與【商品清單】驅動，包含所有「連線登記」、「採購中」、「待配貨」等品項
  * 徹底解決「連線登記」未出現在現場採購總表或配貨清單的問題！
  */
  function getProcurementOrders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    if (!orderSheet) {
      return { success: false, message: '訂單工作表不存在' };
    }

    // 1. 建立商品圖文對照 Map
    const prodMap = {};
    if (prodSheet) {
      const prodValues = prodSheet.getDataRange().getValues();
      for (let p = 1; p < prodValues.length; p++) {
        const pId = String(prodValues[p][0] || '').trim();
        if (pId) {
          let img = String(prodValues[p][7] || '').trim(); // H 欄: 封面主圖
          if (!img || img.startsWith('[') || img.startsWith('{')) {
            try {
              const parsed = JSON.parse(img);
              if (Array.isArray(parsed) && parsed.length > 0) {
                img = typeof parsed[0] === 'string' ? parsed[0] : (parsed[0].url || parsed[0].imageUrl || '');
              }
            } catch (e) { img = ''; }
          }
          if (!img) {
            const allImgsRaw = String(prodValues[p][13] || '').trim(); // N 欄: 所有圖片
            if (allImgsRaw.startsWith('[')) {
              try {
                const parsedAll = JSON.parse(allImgsRaw);
                if (Array.isArray(parsedAll) && parsedAll.length > 0) {
                  img = typeof parsedAll[0] === 'string' ? parsedAll[0] : (parsedAll[0].url || '');
                }
              } catch (e) {}
            }
          }
          prodMap[pId] = {
            name: String(prodValues[p][1] || '').trim(),
            category: String(prodValues[p][2] || '').trim(),
            price: Number(prodValues[p][4]) || 0,
            imageUrl: img
          };
        }
      }
    }

    // 2. 遍歷【訂單明細】所有未結案訂單
    const orderRows = orderSheet.getDataRange().getValues();
    const orders = [];
    const procureGroupMap = {};

    for (let r = 1; r < orderRows.length; r++) {
      const row = orderRows[r];
      const orderId = String(row[0] || '').trim();
      if (!orderId) continue;

      const shipStatus = String(row[18] || '').trim(); // S 欄: 採購出貨狀態
      // 嚴格排除已完成取貨或已結案的歷史訂單
      if (shipStatus.includes('完成取貨') || shipStatus.includes('結案') || shipStatus.includes('已取消')) {
        continue;
      }

      const orderDate = row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : String(row[1] || '');
      const userId = String(row[2] || '').trim();
      const lineName = String(row[3] || '').trim();
      const productId = String(row[4] || '').trim();
      const productName = String(row[5] || '').trim();
      const spec = String(row[6] || '').trim() || '單一規格';
      const qty = Number(row[7]) || 1;
      const unitPrice = Number(row[8]) || 0;
      const subtotal = Number(row[9]) || (unitPrice * qty);
      const recipientName = String(row[12] || '').trim() || String(row[22] || '').trim() || lineName || '顧客';
      const recipientPhone = String(row[13] || '').trim() || String(row[23] || '').trim();
      const note = String(row[15] || '').trim();
      const payStatus = String(row[16] || '').trim();
      const processNote = String(row[19] || '').trim();
      const checkoutId = String(row[21] || '').trim();
      const realName = String(row[22] || '').trim();
      const memberPhone = String(row[23] || '').trim();

      const prodInfo = prodMap[productId] || {};
      const itemImg = prodInfo.imageUrl || '';

      // 判斷該品項採購狀態標記
      let itemStatus = 'pending'; // 預設連線登記 / 待採購
      if (shipStatus.includes('成功') || shipStatus.includes('採買') || shipStatus.includes('已採購') || shipStatus.includes('配貨中')) {
        itemStatus = 'procured';
      } else if (shipStatus.includes('缺貨') || shipStatus.includes('斷貨') || shipStatus.includes('買不到')) {
        itemStatus = 'out_of_stock';
      } else if (shipStatus.includes('已完成結帳') || shipStatus.includes('出貨')) {
        itemStatus = 'procured';
      }

      orders.push({
        orderId: orderId,
        orderDate: orderDate,
        userId: userId,
        lineName: lineName,
        realName: realName,
        name: recipientName,
        phone: recipientPhone || memberPhone,
        productId: productId,
        productName: productName,
        spec: spec,
        quantity: qty,
        unitPrice: unitPrice,
        subtotal: subtotal,
        shippingStatus: shipStatus || '連線登記',
        payStatus: payStatus,
        checkoutId: checkoutId,
        note: note,
        processNote: processNote,
        imageUrl: itemImg,
        status: itemStatus
      });

      // 匯總至 Station 1 現場採購總表
      const groupKey = `${productId}__${spec}`;
      if (!procureGroupMap[groupKey]) {
        procureGroupMap[groupKey] = {
          key: groupKey,
          productId: productId,
          productName: productName,
          spec: spec,
          unitPrice: unitPrice,
          imageUrl: itemImg,
          totalQty: 0,
          procuredQty: 0,
          outOfStockQty: 0,
          pendingQty: 0,
          buyerDetails: []
        };
      }

      procureGroupMap[groupKey].totalQty += qty;
      if (itemStatus === 'procured') {
        procureGroupMap[groupKey].procuredQty += qty;
      } else if (itemStatus === 'out_of_stock') {
        procureGroupMap[groupKey].outOfStockQty += qty;
      } else {
        procureGroupMap[groupKey].pendingQty += qty;
      }

      procureGroupMap[groupKey].buyerDetails.push({
        orderId: orderId,
        buyerName: recipientName,
        lineName: lineName,
        qty: qty,
        status: itemStatus,
        shipStatus: shipStatus
      });
    }

    const procureGroups = Object.values(procureGroupMap);

    return {
      success: true,
      count: orders.length,
      orders: orders,
      procureGroups: procureGroups
    };
  }

  /**
  * 賣家管理者更新包裹與出貨狀態 (包含完成取貨點掉歸檔、還原待出貨、填寫寄件單號)
  */
  function handleUpdateParcelStatus(data) {
    const checkoutId = String(data.checkoutId || '').trim();
    const newStatus = String(data.status || '').trim(); // 例: '已完成取貨' 或 '已完成出貨' 或 '已完成結帳'
    const trackingCode = String(data.trackingCode || '').trim();

    if (!checkoutId) return { success: false, message: '缺少結帳單號 checkoutId' };
    if (!newStatus) return { success: false, message: '缺少欲更新的狀態 status' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    if (!orderSheet) return { success: false, message: '訂單工作表不存在' };

    const rows = orderSheet.getDataRange().getValues();
    let updatedCount = 0;

    for (let i = 1; i < rows.length; i++) {
      const rCheckoutId = String(rows[i][21] || '').trim(); // 第 22 欄 (V 欄)：結帳編號
      if (rCheckoutId === checkoutId) {
        // 第 19 欄 (S 欄)：採購出貨狀態
        orderSheet.getRange(i + 1, 19).setValue(newStatus);
        if (trackingCode) {
          // 第 21 欄 (U 欄)：超商寄件單號 / 宅配託運單號
          orderSheet.getRange(i + 1, 21).setValue(trackingCode);
        }
        updatedCount++;
      }
    }

    return {
      success: true,
      message: `已成功將結帳單【${checkoutId}】(${updatedCount} 筆商品) 狀態更新為【${newStatus}】！`,
      updatedCount: updatedCount,
      checkoutId: checkoutId,
      status: newStatus
    };
  }

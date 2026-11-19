/**
 * ==============================================================================
 * 【05_TelegramService.gs】Telegram 即時營運推播機器人模組
 * ==============================================================================
 * 說明：專門處理 Telegram 訊息排版 (HTML Mode)、訪客稱呼格式化、瀏覽點擊通報與連線測試。
 * ==============================================================================
 */

/**
  * HTML 特殊符號轉義，防止 Telegram HTML 解析模式拋錯
  */
  function escapeHtmlText(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

/**
  * 格式化訪客稱呼 (統稱「訪客稱呼」)
  * 格式規範：LINE暱稱（真實名字/手機號碼）
  * 若未登記真實資訊則僅顯示：LINE暱稱 (若無暱稱則顯示「訪客」)
  */
  function formatVisitorDisplayName(lineName, realName, phone) {
    const lName = String(lineName || '').trim() || '訪客';
    const rName = String(realName || '').trim();
    const pNum = String(phone || '').replace(/[-\s]/g, '').trim();

    const details = [];
    if (rName) details.push(rName);
    if (pNum) details.push(pNum);

    if (details.length > 0) {
      return `${lName}（${details.join('/')}）`;
    }
    return lName;
  }

/**
  * 發送 Telegram 機器人即時推播訊息
  * @param {string} text 訊息內容 (支援 HTML 排版: <b>, <i>, <code>, <a>)
  */
  function sendTelegramMessage(text) {
    try {
      const settings = getRawSystemSettingsInternal();
      const token = settings.TELEGRAM_BOT_TOKEN;
      const chatId = settings.TELEGRAM_CHAT_ID;

      // 若尚未填寫 Token 或 Chat ID，則優雅靜默略過，不影響主業務運行
      if (!token || !chatId) {
        return { success: false, reason: '未設定 TELEGRAM_BOT_TOKEN 或 TELEGRAM_CHAT_ID' };
      }

      const url = `https://api.telegram.org/bot${token}/sendMessage`;
      const payload = {
        chat_id: String(chatId).trim(),
        text: text,
        parse_mode: 'HTML',
        disable_web_page_preview: true
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const resText = response.getContentText() || '{}';
      const result = JSON.parse(resText);

      if (code === 200 && result.ok) {
        return { success: true };
      } else {
        console.warn('Telegram API 回傳異常:', code, resText);
        return { success: false, error: result };
      }
    } catch (err) {
      console.warn('sendTelegramMessage 執行失敗:', err);
      return { success: false, error: err.toString() };
    }
  }

/**
  * 🔔 發送【網址觸擊 / 訪客進站】即時 Telegram 通報至管理者
  * 內建 60 秒 Cache 防灌訊保護機制，即時掌握每一位訪客動態
  */
  function sendPageViewAlertEmail(info) {
    try {
      const pagePath = info.pagePath || '';
      const pageTitle = info.pageTitle || pagePath;
      const productId = info.productId || '-';
      const newViews = info.newViews || 1;
      const userName = String(info.userName || '').trim();
      const userId = String(info.userId || '').trim();
      const now = info.now instanceof Date ? info.now : new Date();
      const nowStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

      // 🔍 智慧身分反查與「進站即自動建立初階顧客歸戶卡」
      let customerRealName = '';
      let customerLineNick = '';
      let customerPhone = '';
      if (userId && !userId.startsWith('GUEST_') && userId !== 'LINE_GUEST' && userId !== 'TEST_USER_999') {
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS || '顧客歸戶');
          if (!custSheet) {
            custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS || '顧客歸戶');
            custSheet.appendRow([
              'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '加入會員時間', 
              '最後進站時間', '最後瀏覽商品/頁面', '進站累計天數', '商品與頁面瀏覽次數', '歷史訂單數', '總消費金額', 
              '會員等級', '顧客生命狀態', '顧客來源', '常用超商取件門市', '黑名單標記', '賣家管理備註/標籤'
            ]);
            custSheet.getRange(1, 1, 1, 17).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
            custSheet.getRange("D:D").setNumberFormat('@');
          }

          ensureOrderAndCustomerSchema(ss, null, custSheet);

          const custData = custSheet.getDataRange().getValues();
          let userFound = false;
          for (let c = 1; c < custData.length; c++) {
            if (String(custData[c][0]).trim() === userId) {
              userFound = true;
              customerLineNick = String(custData[c][1] || '').trim();
              customerRealName = String(custData[c][2] || '').trim();
              customerPhone = String(custData[c][3] || '').trim();
              
              // 若有最新 LINE 暱稱且試算表為空或不同，自動更新暱稱
              if (userName && customerLineNick !== userName) {
                custSheet.getRange(c + 1, 2).setValue(userName);
                customerLineNick = userName;
              }

              // 🌟 1. 最後進站時間 (Col 6)
              const prevVisitRaw = custData[c][5];
              custSheet.getRange(c + 1, 6).setValue(nowStr);

              // 🌟 2. 最後瀏覽商品/頁面 (Col 7)
              custSheet.getRange(c + 1, 7).setValue(pageTitle);

              // 🌟 3. 進站累計天數 (Col 8: 以日計算，當天重複造訪不累加)
              let prevVisitDateStr = '';
              if (prevVisitRaw instanceof Date) {
                prevVisitDateStr = Utilities.formatDate(prevVisitRaw, 'Asia/Taipei', 'yyyy/MM/dd');
              } else if (prevVisitRaw) {
                prevVisitDateStr = String(prevVisitRaw).trim().substring(0, 10).replace(/-/g, '/');
              }
              const todayDateStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd');
              if (prevVisitDateStr !== todayDateStr) {
                const curVisitDays = Number(custData[c][7]) || 0;
                custSheet.getRange(c + 1, 8).setValue(curVisitDays + 1);
              }

              // 🌟 4. 商品與頁面瀏覽次數 (Col 9: 每次點閱造訪皆累加 +1)
              const curPageViews = Number(custData[c][8]) || 0;
              custSheet.getRange(c + 1, 9).setValue(curPageViews + 1);

              // 🌟 5. 顧客生命狀態 (Col 13: 活躍)
              custSheet.getRange(c + 1, 13).setValue('🟢 活躍');
              break;
            }
          }

          // 🌟 核心功能：進站即自動建立初階顧客歸戶卡
          // 若該真實 LINE 好友在試算表中完全無資料，第一時間自動建立初階檔案
          if (!userFound && (userName || userId)) {
            const defaultSource = 'LINE 好友/群組';
            custSheet.appendRow([
              userId,                 // Col 1: LINE_User_ID
              userName || 'LINE 會員', // Col 2: 最新暱稱
              '',                     // Col 3: 真實姓名 (待填寫)
              '',                     // Col 4: 電話 (待填寫)
              nowStr,                 // Col 5: 加入會員時間
              nowStr,                 // Col 6: 最後進站時間
              pageTitle,              // Col 7: 最後瀏覽商品/頁面
              1,                      // Col 8: 進站累計天數 (以日計)
              1,                      // Col 9: 商品與頁面瀏覽次數 (累計PV)
              0,                      // Col 10: 歷史訂單數
              0,                      // Col 11: 總消費金額
              '一般會員',             // Col 12: 會員等級
              '🟢 活躍',              // Col 13: 顧客生命狀態
              defaultSource,          // Col 14: 顧客來源
              '',                     // Col 15: 常用超商取件門市
              '正常',                 // Col 16: 黑名單標記
              ''                      // Col 17: 賣家管理備註/標籤
            ]);
            customerLineNick = userName || 'LINE 會員';
          }

          // 寫入進站逛店動態安全 Log
          logCustomerActivity(
            userId,
            customerRealName ? `${customerRealName} (${customerLineNick})` : customerLineNick,
            '👀 進站逛店',
            `瀏覽: ${pageTitle}${productId !== '-' ? ' (' + productId + ')' : ''}`,
            pagePath,
            info.device || ''
          );
        } catch (custErr) {
          console.warn('反查或自動建立初階顧客歸戶失敗:', custErr);
        }
      }

      // 🌟 彙整訪客稱呼：嚴格遵循「LINE暱稱（真實名字/手機號碼）」規範
      const lineNick = customerLineNick || userName;
      const visitorDisplay = formatVisitorDisplayName(lineNick, customerRealName, customerPhone);

      // 🛡️ 防短時間灌訊保護 (以頁面與訪客身分為 key，60 秒內不重複推播相同訪客點擊相同頁面)
      const tgVisitorAlertCacheKey = userId || lineNick || 'anon';
      const cacheKey = 'pv_tg_alert_' + encodeURIComponent(pagePath) + '_' + encodeURIComponent(tgVisitorAlertCacheKey);
      const cache = CacheService.getScriptCache();
      if (cache.get(cacheKey)) {
        return; // 60 秒冷卻略過
      }
      cache.put(cacheKey, '1', 60);

      // 🔔 發送【訪客瀏覽通報】Telegram 通知
      const rawSettings = getRawSystemSettingsInternal();
      if (rawSettings.TELEGRAM_NOTIFY_PAGE_VIEWS !== 'NO') {
        const prodPart = (productId && productId !== '-') ? `• <b>商品識別</b>：<code>${escapeHtmlText(productId)}</code>\n` : '';
        const tgMsg = 
          `👁️ <b>【訪客瀏覽通報】</b>\n` +
          `• <b>訪客稱呼</b>：${escapeHtmlText(visitorDisplay)}\n` +
          `• <b>瀏覽頁面</b>：${escapeHtmlText(pageTitle)}\n` +
          prodPart +
          `• <b>累計次數</b>：第 ${newViews} 次點擊\n` +
          `• <b>使用裝置</b>：${escapeHtmlText(info.device || '行動/電腦裝置')}\n` +
          `• <b>進站時間</b>：${nowStr}`;

        sendTelegramMessage(tgMsg);
      }
    } catch (err) {
      console.warn('sendPageViewAlertNotification 執行失敗:', err);
    }
  }

/**
 * 試算表選單專用：強制觸發授權並測試發送 Telegram 訊息
 */
function forceAuthAndTest() {
  const result = sendTelegramMessage('🔔 <b>【連線系統測試】</b>\n恭喜！您的 Telegram 機器人已成功與 Google Apps Script 建立即時連線！\n• <b>訪客稱呼</b>：系統管理員\n• <b>連線時間</b>：' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'));
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('🎉 授權與連線成功！', 'Telegram 機器人已連線成功！\n請打開您的手機或電腦版 Telegram 查看測試推播訊息！', ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ 尚未收到訊息', '請確認試算表【系統設定】中的 TELEGRAM_BOT_TOKEN 與 TELEGRAM_CHAT_ID 是否已填寫正確。\n詳細回應：' + JSON.stringify(result), ui.ButtonSet.OK);
  }
}

/**
 * 強制觸發 Google 外部網路 (UrlFetchApp) 授權彈窗的專用函式 (不包 try-catch，強制 Google 彈出授權視窗)
 */
function forceAuth() {
  UrlFetchApp.fetch('https://api.telegram.org');
}

/**
 * 專供 Apps Script 編輯器手動執行「一鍵觸發授權」與「連線測試」的專屬函式
 */
function testTelegramInEditor() {
  const result = sendTelegramMessage('🔔 <b>【連線系統測試】</b>\n恭喜！您的 Telegram 機器人已成功與 Google Apps Script 建立即時連線！\n• <b>訪客稱呼</b>：系統管理員\n• <b>連線時間</b>：' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'));
  Logger.log(JSON.stringify(result));
  return result;
}

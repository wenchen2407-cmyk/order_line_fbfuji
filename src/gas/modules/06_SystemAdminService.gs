/**
 * ==============================================================================
 * 【06_SystemAdminService.gs】試算表結構維護與管理設定模組
 * ==============================================================================
 * 說明：專門處理試算表初始化 (setupSpreadsheet)、自動無損擴充欄位、系統設定讀取與瀏覽統計。
 * ==============================================================================
 */

/**
  * 試算表初次安裝設定與欄位升級：自動補齊成本價、重量備註與多圖欄位
  * 可在 Apps Script 編輯器中直接執行此函式升級試算表結構
  */
  function setupSpreadsheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 商品清單工作表 (升級支援原幣稅前金額、稅金、匯率、重量、公式連動欄位與最後更新時間)
    let prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    const prodHeaders = [
      '商品編號', '商品名稱', '供應模式', '專櫃原價', '連線代購價', '現貨庫存', 
      '規格清單(JSON)', '封面主圖網址', '商品描述', '狀態', '建立時間',
      '成本價(NT$)', '採購備註說明', '所有圖片清單(JSON)', '指定結單時間',
      '採購幣別', '原幣稅前金額', '稅金(原幣)', '採購匯率', '商品重量(g)', '國際運費(元/kg)',
      '最後更新時間'
    ];

    if (!prodSheet) {
      prodSheet = ss.insertSheet(SHEET_NAMES.PRODUCTS);
      prodSheet.appendRow(prodHeaders);
      prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
      
      // 加入範例資料 (成本價採用精確分開四捨五入公式連動：ROUND((Q2+R2)*S2) + ROUND((T2/1000)*U2))
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      prodSheet.appendRow([
        'J2609270001', '日本代購限定 輕量防潑水後背包', '✈️ 連線', 1880, 1450, 999999,
        JSON.stringify(['米白色', '經典黑', '海軍藍']),
        'https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80',
        '日本專櫃直購，輕盈大容量，防潑水尼龍材質！', '上架中', new Date(),
        '=IF(OR(N(Q2)>0, N(T2)>0), ROUND((N(Q2) + N(R2)) * N(S2)) + ROUND((N(T2) / 1000) * IF(N(U2)>0, N(U2), 240)), 850)', '日本專櫃進貨', JSON.stringify(['https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=800&auto=format&fit=crop&q=80']),
        '2026/10/10 23:59',
        'JPY', 3000, 300, 0.22, 380, 240,
        nowStr
      ]);
    } else {
      // 既有表格：升級表頭為最新 22 欄規格 (C 欄直接改為「供應模式」)
      const curCols = prodSheet.getMaxColumns();
      if (curCols < prodHeaders.length) {
        prodSheet.insertColumnsAfter(curCols, prodHeaders.length - curCols);
      }
      prodSheet.getRange(1, 1, 1, prodHeaders.length).setValues([prodHeaders]);
      prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
      
      // 🌟 一鍵同步升級：自動掃描既有商品，補齊獨立原幣/稅金/重量欄位與最後更新時間，並將成本價全面升級為連動公式
      upgradeExistingProductsToFormula(prodSheet);
    }
    // 將 O 欄 (第15欄) 設為純文字格式，避免日期被 Excel/Sheets 自動轉換失真
    prodSheet.getRange("O:O").setNumberFormat('@');

    // 1.1 為【商品清單】C 欄建立「供應模式」下拉選單 (現貨 / 預購 / 連線)
    try {
      const modeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['✈️ 親飛連線', '⚡ 現貨', '📦 預購', '✈️ 連線', '親飛連線', '現貨', '預購', '連線'], true)
        .setAllowInvalid(true)
        .build();
      prodSheet.getRange("C2:C2000").setDataValidation(modeRule);
    } catch (e) {
      console.warn('建立商品清單供應模式下拉選單略過:', e);
    }
    prodSheet.getRange("O:O").setNumberFormat('@');

    // 2. 訂單明細工作表 (升級為 24 欄規格：徹底分離 M/N 欄「結帳收件人姓名/電話」與 W/X 欄「LINE 會員真實姓名/電話」)
    let orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const orderHeaders = [
      '訂單編號', '下單時間', 'LINE_User_ID', 'LINE暱稱', '商品編號', 
      '商品名稱', '選購規格', '數量', '單價', '商品小計', '運費', 
      '訂單總額', '收件人姓名', '收件人電話', '取件方式與地址', 
      '買家備註', '付款狀態', '匯款後五碼', '採購/出貨狀態', '處理備註',
      '包裹追蹤編號', '結帳編號', '真實姓名', '會員電話'
    ];
    if (!orderSheet) {
      orderSheet = ss.insertSheet(SHEET_NAMES.ORDERS);
      orderSheet.appendRow(orderHeaders);
    } else {
      const curOrderCols = orderSheet.getMaxColumns();
      if (curOrderCols < orderHeaders.length) {
        orderSheet.insertColumnsAfter(curOrderCols, orderHeaders.length - curOrderCols);
      }
      orderSheet.getRange(1, 1, 1, orderHeaders.length).setValues([orderHeaders]);
    }
    orderSheet.getRange(1, 1, 1, orderHeaders.length).setBackground('#107c41').setFontColor('#ffffff').setFontWeight('bold');
    orderSheet.getRange("N:N").setNumberFormat('@'); // 收件人電話純文字格式
    orderSheet.getRange("R:R").setNumberFormat('@');
    orderSheet.getRange("U:U").setNumberFormat('@'); // 包裹單號純文字格式防止開頭 0 丟失
    orderSheet.getRange("V:V").setNumberFormat('@'); // 結帳編號 (Checkout_ID) 純文字格式
    orderSheet.getRange("X:X").setNumberFormat('@'); // 會員電話純文字格式防止開頭 0 丟失

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

    // 3. 顧客歸戶表 (升級為 17 欄規格：完整記錄會員動態、進站天數、商品頁面瀏覽次數、消費、選單設定)
    const custHeaders = [
      'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '加入會員時間', 
      '最後進站時間', '最後瀏覽商品/頁面', '進站累計天數', '商品與頁面瀏覽次數', '歷史訂單數', '總消費金額', 
      '會員等級', '顧客生命狀態', '顧客來源', '常用超商取件門市', '黑名單標記', '賣家管理備註/標籤'
    ];
    let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    if (!custSheet) {
      custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
      custSheet.appendRow(custHeaders);
    } else {
      const curCustCols = custSheet.getMaxColumns();
      if (curCustCols < custHeaders.length) {
        custSheet.insertColumnsAfter(curCustCols, custHeaders.length - curCustCols);
      }
      custSheet.getRange(1, 1, 1, custHeaders.length).setValues([custHeaders]);
    }
    custSheet.getRange(1, 1, 1, custHeaders.length).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
    custSheet.getRange("D:D").setNumberFormat('@'); // 電話純文字格式防止開頭 0 丟失
    custSheet.getRange("E:E").setNumberFormat('yyyy/MM/dd HH:mm:ss'); // 加入會員時間
    custSheet.getRange("F:F").setNumberFormat('yyyy/MM/dd HH:mm:ss'); // 最後進站時間
    custSheet.getRange("H:H").setNumberFormat('#,##0'); // 進站累計天數 (以日計)
    custSheet.getRange("I:I").setNumberFormat('#,##0'); // 商品與頁面瀏覽次數 (累計PV)
    custSheet.getRange("J:J").setNumberFormat('#,##0'); // 歷史訂單數
    custSheet.getRange("K:K").setNumberFormat('NT$ #,##0'); // 總消費金額

    // 3.1 為【顧客歸戶】建立下拉選單 (會員等級、顧客生命狀態、顧客來源、黑名單標記)
    try {
      // L 欄 (第12欄): 會員等級
      const tierRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['一般會員', 'VIP', 'VVIP', '批發夥伴'], true)
        .setAllowInvalid(true)
        .build();
      custSheet.getRange("L2:L2000").setDataValidation(tierRule);

      // M 欄 (第13欄): 顧客生命狀態
      const lifeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['🟢 活躍', '🟡 猶豫', '💤 沉睡'], true)
        .setAllowInvalid(true)
        .build();
      custSheet.getRange("M2:M2000").setDataValidation(lifeRule);

      // N 欄 (第14欄): 顧客來源
      const sourceRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['LINE 好友/群組', 'Facebook 社團/粉專', 'Instagram', '親友推薦', '其他'], true)
        .setAllowInvalid(true)
        .build();
      custSheet.getRange("N2:N2000").setDataValidation(sourceRule);

      // P 欄 (第16欄): 黑名單標記
      const blacklistRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['正常', '棄單觀察', '黑名單拒絕往來'], true)
        .setAllowInvalid(true)
        .build();
      custSheet.getRange("P2:P2000").setDataValidation(blacklistRule);
    } catch (valErr) {
      console.warn('建立顧客表下拉選單略過:', valErr);
    }

    // 3.1 會員動態歷程 Log (安全稽核流水帳：不可竄改之全生命週期動態軌跡)
    const logHeaders = ['記錄時間', 'LINE_User_ID', '會員真實姓名/暱稱', '事件類型', '動態詳情', '來源頁面/管道', '裝置環境'];
    let logSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
      logSheet.appendRow(logHeaders);
    }
    logSheet.getRange(1, 1, 1, logHeaders.length).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
    logSheet.getRange("A:A").setNumberFormat('yyyy/MM/dd HH:mm:ss');
    logSheet.getRange("B:B").setNumberFormat('@');

    // 🌟 自動遷移與校準既有訂單明細與顧客歸戶資料（徹底分離「會員真實姓名/電話」與「結帳收件人姓名/電話」）
    try {
      upgradeOrderAndCustomerSheets(ss);
    } catch (migErr) {
      console.warn('遷移訂單與顧客收件人欄位警告:', migErr);
    }

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
      ['COD_MART_NOTICE', '整理完畢後，賣家將於群組或私訊發送專屬賣貨便/好賣+賣場連結供您下單出貨！', '賣貨便/好賣+ 提示說明'],
      ['ADMIN_LINE_UIDS', 'Uee7765ed3c1819cb94fd404cfab8553b', '管理者 LINE User ID 白名單 (多筆以逗號或換行分隔)'],
      ['ADMIN_API_KEY', 'admin8888', '管理後台安全通行碼 (PIN 碼)，忘記時可直接在此欄位查看或修改'],
      ['TELEGRAM_BOT_TOKEN', '', 'Telegram 機器人 Token (由 @BotFather 取得，留空則不推播)'],
      ['TELEGRAM_CHAT_ID', '', 'Telegram 接收通報的 Chat ID (個人 ID 或群組 ID)'],
      ['TELEGRAM_NOTIFY_PAGE_VIEWS', 'YES', '是否將進站與商品瀏覽通報傳送到 Telegram (YES/NO)'],
      ['TELEGRAM_NOTIFY_ORDERS', 'YES', '是否將買家新下單/加單登記即時傳送到 Telegram (YES/NO)'],
      ['TELEGRAM_NOTIFY_PAYMENTS', 'YES', '是否將買家回報匯款/出貨結帳即時傳送到 Telegram (YES/NO)']
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

    try {
      if (typeof setupDailyLogCleanerTrigger === 'function') {
        setupDailyLogCleanerTrigger();
      }
    } catch (trigErr) {
      console.warn('設置每日日誌清理排程略過:', trigErr);
    }

    return '工作表與系統設定更新升級完成！';
  }

/**
  * 🌟 確保【訂單明細】(24 欄) 與【顧客歸戶】(13 欄) 表頭與欄數已升級，若尚未升級則自動觸發無損升級與資料遷移
  */
  function ensureOrderAndCustomerSchema(ss, orderSheet, custSheet) {
    try {
      const activeSs = ss || SpreadsheetApp.getActiveSpreadsheet();
      const oSheet = orderSheet || activeSs.getSheetByName(SHEET_NAMES.ORDERS);
      const cSheet = custSheet || activeSs.getSheetByName(SHEET_NAMES.CUSTOMERS);
      const lSheet = activeSs.getSheetByName(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');

      let needsUpgrade = false;
      if (oSheet) {
        const oCols = oSheet.getLastColumn();
        if (oCols < 24) {
          needsUpgrade = true;
        } else {
          const wHeader = String(oSheet.getRange(1, 23).getValue() || '').trim();
          const nHeader = String(oSheet.getRange(1, 14).getValue() || '').trim();
          if (wHeader !== '真實姓名' || nHeader !== '收件人電話') {
            needsUpgrade = true;
          }
        }
      }
      if (cSheet && !needsUpgrade) {
        const cCols = cSheet.getLastColumn();
        if (cCols < 17) {
          needsUpgrade = true;
        } else {
          const eHeader = String(cSheet.getRange(1, 5).getValue() || '').trim();
          const fHeader = String(cSheet.getRange(1, 6).getValue() || '').trim();
          const hHeader = String(cSheet.getRange(1, 8).getValue() || '').trim();
          const iHeader = String(cSheet.getRange(1, 9).getValue() || '').trim();
          if (eHeader !== '加入會員時間' || fHeader !== '最後進站時間' || !hHeader.includes('天數') || iHeader !== '商品與頁面瀏覽次數') {
            needsUpgrade = true;
          }
        }
      }
      if (!lSheet) {
        needsUpgrade = true;
      }

      if (needsUpgrade) {
        upgradeOrderAndCustomerSheets(activeSs);
      }
    } catch (err) {
      console.warn('自動檢查訂單與顧客表結構略過:', err);
    }
  }

/**
  * 🌟 自動升級與校準【訂單明細】與【顧客歸戶】欄位結構，徹底分離「LINE 帳號登記真實姓名/電話」與「結帳收件人姓名/電話」
  * 可直接在 Apps Script 編輯器選取此函式執行，或由 setupSpreadsheet() / API 自動觸發
  */
  function upgradeOrderAndCustomerSheets(targetSs) {
    const ss = targetSs || SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);

    const orderHeaders = [
      '訂單編號', '下單時間', 'LINE_User_ID', 'LINE暱稱', '商品編號', 
      '商品名稱', '選購規格', '數量', '單價', '商品小計', '運費', 
      '訂單總額', '收件人姓名', '收件人電話', '取件方式與地址', 
      '買家備註', '付款狀態', '匯款後五碼', '採購/出貨狀態', '處理備註',
      '包裹追蹤編號', '結帳編號', '真實姓名', '會員電話'
    ];

    const logHeaders = ['記錄時間', 'LINE_User_ID', '會員真實姓名/暱稱', '事件類型', '動態詳情', '來源頁面/管道', '裝置環境'];

    // 1. 確保【會員動態Log】安全稽核軌跡表存在
    let logSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_NAMES.CUSTOMER_LOGS || '會員動態Log');
      logSheet.appendRow(logHeaders);
    }
    logSheet.getRange(1, 1, 1, logHeaders.length).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
    logSheet.getRange("A:A").setNumberFormat('yyyy/MM/dd HH:mm:ss');
    logSheet.getRange("B:B").setNumberFormat('@');

    // 2. 升級【訂單明細】為 24 欄
    if (orderSheet) {
      const maxOrderCols = orderSheet.getMaxColumns();
      if (maxOrderCols < orderHeaders.length) {
        orderSheet.insertColumnsAfter(maxOrderCols, orderHeaders.length - maxOrderCols);
      }
      orderSheet.getRange(1, 1, 1, orderHeaders.length).setValues([orderHeaders]);
      orderSheet.getRange(1, 1, 1, orderHeaders.length).setBackground('#107c41').setFontColor('#ffffff').setFontWeight('bold');
      orderSheet.getRange("N:N").setNumberFormat('@');
      orderSheet.getRange("R:R").setNumberFormat('@');
      orderSheet.getRange("U:U").setNumberFormat('@');
      orderSheet.getRange("V:V").setNumberFormat('@');
      orderSheet.getRange("X:X").setNumberFormat('@');
    }

    // 3. 升級【顧客歸戶】為 17 欄 (無損平滑資料遷移：進站天數、商品瀏覽次數、下拉選單)
    const custHeaders = [
      'LINE_User_ID', '最新暱稱', '真實姓名', '電話', '加入會員時間', 
      '最後進站時間', '最後瀏覽商品/頁面', '進站累計天數', '商品與頁面瀏覽次數', '歷史訂單數', '總消費金額', 
      '會員等級', '顧客生命狀態', '顧客來源', '常用超商取件門市', '黑名單標記', '賣家管理備註/標籤'
    ];

    if (custSheet) {
      const nowStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
      const custLastRow = custSheet.getLastRow();
      const custLastCol = custSheet.getLastColumn();

      let rawHeaders = [];
      if (custLastCol > 0) {
        rawHeaders = custSheet.getRange(1, 1, 1, custLastCol).getValues()[0].map(h => String(h || '').trim());
      }

      const isAlready17 = custLastCol >= 17 && rawHeaders[4] === '加入會員時間' && rawHeaders[5] === '最後進站時間' && rawHeaders[8] === '商品與頁面瀏覽次數';

      if (!isAlready17 && custLastRow > 1) {
        const is16Col = (custLastCol === 16 || rawHeaders.length === 16) && rawHeaders[4] === '加入會員時間' && rawHeaders[5] === '最後進站時間';
        const oldValues = custSheet.getRange(2, 1, custLastRow - 1, custLastCol).getValues();
        const migratedRows = [];

        if (is16Col) {
          // 從 16 欄升級至 17 欄：在第 8 欄 (進站累計次數) 後方插入「商品與頁面瀏覽次數」
          for (let r = 0; r < oldValues.length; r++) {
            const row = oldValues[r];
            const uid = String(row[0] || '').trim();
            if (!uid) continue;
            const visitDays = Number(row[7]) || 1;
            const pageViews = Math.max(visitDays, Number(row[8]) > 0 ? Number(row[8]) : visitDays);
            migratedRows.push([
              row[0], row[1], row[2], formatPhoneAsText(row[3]), row[4],
              row[5], row[6], visitDays, pageViews,
              Number(row[8]) || 0,   // 歷史訂單數
              Number(row[9]) || 0,   // 總消費金額
              row[10] || '一般會員', // 會員等級
              row[11] || '🟢 活躍',  // 顧客生命狀態
              row[12] || '',         // 顧客來源
              row[13] || '',         // 常用超商取件門市
              row[14] || '正常',     // 黑名單標記
              row[15] || ''          // 賣家管理備註/標籤
            ]);
          }
        } else {
          // 從舊版 13 欄無損遷移至 17 欄全新架構
          const idxUid = rawHeaders.indexOf('LINE_User_ID');
          const idxNick = rawHeaders.indexOf('最新暱稱');
          const idxRealName = rawHeaders.indexOf('真實姓名');
          const idxPhone = rawHeaders.indexOf('電話');
          const idxOldAddr = rawHeaders.indexOf('常用寄送地址');
          const idxOrderCount = rawHeaders.indexOf('歷史訂單數');
          const idxSpend = rawHeaders.indexOf('總消費金額');
          const idxFirstBuy = rawHeaders.indexOf('首購日期');
          const idxLastOrder = rawHeaders.indexOf('最後下單日期');
          const idxBlacklist = rawHeaders.indexOf('黑名單標記');
          const idxSource = rawHeaders.indexOf('顧客來源');
          const idxRecipName = rawHeaders.indexOf('最近收件人姓名');
          const idxRecipPhone = rawHeaders.indexOf('最近收件人電話');

          for (let r = 0; r < oldValues.length; r++) {
            const row = oldValues[r];
            const uid = idxUid !== -1 ? String(row[idxUid] || '').trim() : String(row[0] || '').trim();
            if (!uid) continue;

            const nick = idxNick !== -1 ? String(row[idxNick] || '').trim() : String(row[1] || '').trim();
            const realName = idxRealName !== -1 ? String(row[idxRealName] || '').trim() : String(row[2] || '').trim();
            const phone = idxPhone !== -1 ? formatPhoneAsText(row[idxPhone]) : formatPhoneAsText(row[3]);
            const firstBuy = idxFirstBuy !== -1 ? row[idxFirstBuy] : '';
            const lastOrder = idxLastOrder !== -1 ? row[idxLastOrder] : '';
            const joinedAt = firstBuy || lastOrder || nowStr;
            const lastVisitAt = lastOrder || firstBuy || nowStr;
            const lastViewed = '';
            const orderCount = idxOrderCount !== -1 ? (Number(row[idxOrderCount]) || 0) : (Number(row[5]) || 0);
            const spend = idxSpend !== -1 ? (Number(row[idxSpend]) || 0) : (Number(row[6]) || 0);
            const visitDays = Math.max(1, orderCount > 0 ? orderCount : 1);
            const pageViews = Math.max(visitDays, orderCount > 0 ? orderCount * 3 : 1);
            const tier = spend >= 10000 ? 'VVIP' : (spend >= 3000 ? 'VIP' : '一般會員');
            const lifeStatus = '🟢 活躍';
            const source = idxSource !== -1 ? String(row[idxSource] || '') : (String(row[10] || ''));
            const oldAddr = idxOldAddr !== -1 ? String(row[idxOldAddr] || '') : (String(row[4] || ''));
            const recipName = idxRecipName !== -1 ? String(row[idxRecipName] || '') : (String(row[11] || ''));
            const recipPhone = idxRecipPhone !== -1 ? String(row[idxRecipPhone] || '') : (String(row[12] || ''));
            const defaultStore = oldAddr || (recipName ? `${recipName} ${recipPhone}`.trim() : '');
            const blacklist = idxBlacklist !== -1 ? (String(row[idxBlacklist] || '正常')) : '正常';
            const notes = '';

            migratedRows.push([
              uid, nick, realName, phone, joinedAt, lastVisitAt, lastViewed, visitDays, pageViews,
              orderCount, spend, tier, lifeStatus, source, defaultStore, blacklist, notes
            ]);
          }
        }

        custSheet.clearContents();
        const curCols = custSheet.getMaxColumns();
        if (curCols < custHeaders.length) {
          custSheet.insertColumnsAfter(curCols, custHeaders.length - curCols);
        }
        custSheet.getRange(1, 1, 1, custHeaders.length).setValues([custHeaders]);
        if (migratedRows.length > 0) {
          custSheet.getRange(2, 1, migratedRows.length, custHeaders.length).setValues(migratedRows);
        }
      } else {
        const curCols = custSheet.getMaxColumns();
        if (curCols < custHeaders.length) {
          custSheet.insertColumnsAfter(curCols, custHeaders.length - curCols);
        }
        custSheet.getRange(1, 1, 1, custHeaders.length).setValues([custHeaders]);
      }

      custSheet.getRange(1, 1, 1, custHeaders.length).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
      custSheet.getRange("D:D").setNumberFormat('@'); // 電話純文字格式
      custSheet.getRange("E:E").setNumberFormat('yyyy/MM/dd HH:mm:ss'); // 加入會員時間
      custSheet.getRange("F:F").setNumberFormat('yyyy/MM/dd HH:mm:ss'); // 最後進站時間
      custSheet.getRange("H:H").setNumberFormat('#,##0'); // 進站累計天數
      custSheet.getRange("I:I").setNumberFormat('#,##0'); // 商品與頁面瀏覽次數
      custSheet.getRange("J:J").setNumberFormat('#,##0'); // 歷史訂單數
      custSheet.getRange("K:K").setNumberFormat('NT$ #,##0'); // 總消費金額

      // 建立下拉選單驗證
      try {
        const tierRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['一般會員', 'VIP', 'VVIP', '批發夥伴'], true)
          .setAllowInvalid(true).build();
        custSheet.getRange("L2:L2000").setDataValidation(tierRule);

        const lifeRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['🟢 活躍', '🟡 猶豫', '💤 沉睡'], true)
          .setAllowInvalid(true).build();
        custSheet.getRange("M2:M2000").setDataValidation(lifeRule);

        const sourceRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['LINE 好友/群組', 'Facebook 社團/粉專', 'Instagram', '親友推薦', '其他'], true)
          .setAllowInvalid(true).build();
        custSheet.getRange("N2:N2000").setDataValidation(sourceRule);

        const blacklistRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['正常', '棄單觀察', '黑名單拒絕往來'], true)
          .setAllowInvalid(true).build();
        custSheet.getRange("P2:P2000").setDataValidation(blacklistRule);
      } catch (valErr) {
        console.warn('建立下拉選單略過:', valErr);
      }
    }

    if (!orderSheet || !custSheet) {
      return '表頭升級完成';
    }

    const custLastRow = custSheet.getLastRow();
    const orderLastRow = orderSheet.getLastRow();

    // 讀取顧客表資料，建立 uid -> 會員真實姓名與電話映射
    const custMap = {};
    let custValues = [];
    if (custLastRow > 1) {
      custValues = custSheet.getRange(2, 1, custLastRow - 1, 17).getValues();
      for (let c = 0; c < custValues.length; c++) {
        const uid = String(custValues[c][0] || '').trim();
        if (uid) {
          custMap[uid] = {
            rowIdx: c,
            realName: String(custValues[c][2] || '').trim(),
            phone: String(custValues[c][3] || '').trim(),
            defaultStore: String(custValues[c][14] || '').trim()
          };
        }
      }
    }

    // 記錄每位顧客最近一次「已結帳」的包裹收件人資訊與未結帳時的會員本人姓名
    const latestCheckoutRecipByUid = {};
    const inferredMemberByUid = {};

    if (orderLastRow > 1) {
      const orderRange = orderSheet.getRange(2, 1, orderLastRow - 1, 24);
      const orderValues = orderRange.getValues();
      let orderModified = false;

      for (let i = 0; i < orderValues.length; i++) {
        const row = orderValues[i];
        const uid = String(row[2] || '').trim();
        const colRecipName = String(row[12] || '').trim();  // M 欄 (13): 收件人姓名
        const colRecipPhone = String(row[13] || '').trim(); // N 欄 (14): 收件人電話
        const colDelivery = String(row[14] || '').trim();   // O 欄 (15): 取件方式與地址
        const payStatus = String(row[16] || '').trim();     // Q 欄 (17): 付款狀態
        const shipStatus = String(row[18] || '').trim();    // S 欄 (19): 採購/出貨狀態
        const checkoutId = String(row[21] || '').trim();    // V 欄 (22): 結帳編號
        let colMemberName = String(row[22] || '').trim();   // W 欄 (23): 真實姓名 (訂購會員)
        let colMemberPhone = String(row[23] || '').trim();  // X 欄 (24): 會員電話 (訂購會員)

        const isCheckedOut = Boolean(checkoutId) ||
                            shipStatus.includes('已完成結帳') ||
                            shipStatus.includes('出貨') ||
                            shipStatus.includes('完成取貨') ||
                            payStatus.includes('待付款') ||
                            payStatus.includes('貨到付款') ||
                            payStatus.includes('對帳') ||
                            (Boolean(colDelivery) && colDelivery !== '[待出貨結帳填寫]');

        if (!isCheckedOut) {
          if (!colMemberName && colRecipName) {
            colMemberName = colRecipName;
            row[22] = colMemberName;
            orderModified = true;
          }
          if (!colMemberPhone && colRecipPhone) {
            colMemberPhone = colRecipPhone;
            row[23] = formatPhoneAsText(colMemberPhone);
            orderModified = true;
          }
          if (colRecipName !== '' || colRecipPhone !== '') {
            row[12] = '';
            row[13] = '';
            orderModified = true;
          }
          if (uid && (colMemberName || colMemberPhone)) {
            inferredMemberByUid[uid] = {
              realName: colMemberName,
              phone: colMemberPhone
            };
          }
        } else {
          if (colRecipPhone && !String(row[13]).startsWith("'")) {
            row[13] = formatPhoneAsText(colRecipPhone);
            orderModified = true;
          }
          if (uid && (colRecipName || colRecipPhone)) {
            latestCheckoutRecipByUid[uid] = {
              recipientName: colRecipName,
              recipientPhone: colRecipPhone,
              deliveryAddress: (colDelivery && colDelivery !== '[待出貨結帳填寫]') ? colDelivery : ''
            };
          }
        }

        // 🌟 將【顧客歸戶】中的會員本人「真實姓名 (C欄)」與「電話 (D欄)」權威同步至【訂單明細】W 欄與 X 欄
        const custInfo = uid ? custMap[uid] : null;
        if (custInfo) {
          if (custInfo.realName && colMemberName !== custInfo.realName) {
            row[22] = custInfo.realName;
            colMemberName = custInfo.realName;
            orderModified = true;
          }
          if (custInfo.phone) {
            const formattedCustPhone = formatPhoneAsText(custInfo.phone);
            if (String(row[23] || '').replace(/^'/, '').trim() !== custInfo.phone.replace(/^'/, '').trim()) {
              row[23] = formattedCustPhone;
              colMemberPhone = custInfo.phone;
              orderModified = true;
            }
          }
        } else if (colMemberPhone && !String(row[23]).startsWith("'")) {
          row[23] = formatPhoneAsText(colMemberPhone);
          orderModified = true;
        }
      }

      if (orderModified) {
        orderRange.setValues(orderValues);
        orderSheet.getRange("N:N").setNumberFormat('@');
        orderSheet.getRange("X:X").setNumberFormat('@');
      }
    }

    // 4. 將已結帳包裹之門市/地址回填至【顧客歸戶】常用超商取件門市 (第 15 欄 O)
    if (custLastRow > 1 && custValues.length > 0) {
      let custModified = false;
      for (let c = 0; c < custValues.length; c++) {
        const uid = String(custValues[c][0] || '').trim();
        if (!uid) continue;

        const latestRecip = latestCheckoutRecipByUid[uid];
        if (latestRecip && latestRecip.deliveryAddress) {
          if (!String(custValues[c][14] || '').trim()) {
            custValues[c][14] = latestRecip.deliveryAddress; // 第 15 欄 (O): 常用超商取件門市
            custModified = true;
          }
        }

        // 若顧客歸戶 C/D 欄為空，且有未結帳訂單推導出的會員姓名電話，則補齊
        const inferred = inferredMemberByUid[uid];
        if (inferred) {
          if (!String(custValues[c][2] || '').trim() && inferred.realName) {
            custValues[c][2] = inferred.realName;
            custModified = true;
          }
          if (!String(custValues[c][3] || '').trim() && inferred.phone) {
            custValues[c][3] = formatPhoneAsText(inferred.phone);
            custModified = true;
          }
        }
        if (custValues[c][3] && !String(custValues[c][3]).startsWith("'")) {
          custValues[c][3] = formatPhoneAsText(custValues[c][3]);
          custModified = true;
        }
      }

      if (custModified) {
        custSheet.getRange(2, 1, custValues.length, 17).setValues(custValues);
        custSheet.getRange("D:D").setNumberFormat('@');
      }
    }

    // 5. 寫入升級成功安全 Log
    logCustomerActivity('SYSTEM', '系統維護', '🔄 結構升級', '成功升級【顧客歸戶】為 17 欄全新架構 (含進站累計天數、商品瀏覽次數與下拉選單)，並建立【會員動態Log】安全稽核軌跡表', 'upgradeOrderAndCustomerSheets', 'GAS Server');

    SpreadsheetApp.flush();
    return '✅ 已完成【訂單明細】(24欄)、【顧客歸戶】(17欄升級) 與【會員動態Log】安全歷程表建立！';
  }

/**
  * 取得完整系統設定 (僅供後端 GAS 內部安全呼叫，包含金鑰與通行碼)
  */
  function getRawSystemSettingsInternal() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
      if (!sheet) return {};

      const rows = sheet.getDataRange().getValues();
      const settings = {};
      for (let i = 1; i < rows.length; i++) {
        const key = String(rows[i][0] || '').trim();
        const val = rows[i][1];
        if (key) {
          settings[key] = (val !== undefined && val !== null) ? String(val).trim() : '';
        }
      }
      return settings;
    } catch (e) {
      console.warn('getRawSystemSettingsInternal 執行失敗:', e);
      return {};
    }
  }

/**
  * 取得系統設定 (前台買家/公開查詢 API，嚴格排除敏感密鑰)
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
        const trimKey = String(key).trim();
        // 🛡️ 嚴格排除後台管理員白名單 UID、安全通行碼 (PIN) 與 Telegram 機密金鑰，禁止透過公開 API 外洩
        if (trimKey === 'ADMIN_LINE_UIDS' || trimKey === 'ADMIN_API_KEY' || trimKey === 'TELEGRAM_BOT_TOKEN' || trimKey === 'TELEGRAM_CHAT_ID') continue;
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
  * 處理頁面與商品瀏覽人次累計 (非阻塞輕量寫入，絕不佔用全域排隊鎖，確保買家下單 100% 暢通)
  */
  function handleRecordPageView(data) {
    if (!data || !data.page) {
      return { success: false, message: '無效的頁面資料' };
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
      let pageTitle = String(data.title || pagePath).trim();
      const productId = String(data.productId || '-').trim();

      // 🛡️ 關鍵防護 1：若為商品且傳入通用未就緒標題（例如「連線代購即時訂購單」或「商品內頁」），自動反查商品表中的真實商品名稱
      if (productId && productId !== '-') {
        const isGeneric = pageTitle.includes('連線代購即時訂購單') || pageTitle === '商品內頁' || !pageTitle;
        if (isGeneric) {
          const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
          if (prodSheet) {
            const pValues = prodSheet.getDataRange().getValues();
            for (let pIdx = 1; pIdx < pValues.length; pIdx++) {
              if (String(pValues[pIdx][0] || '').trim() === productId) {
                const actualName = String(pValues[pIdx][1] || '').trim();
                if (actualName) {
                  pageTitle = `【商品】${actualName}`;
                  break;
                }
              }
            }
          }
        }
      }

      // 僅排除「賣家管理中心」，買家查單與商品瀏覽皆正常統計並寄發觸擊通報
      const pLower = pagePath.toLowerCase();
      if (pLower === 'admin-card-generator.html' || pageTitle.includes('賣家管理') || pageTitle.includes('發卡推播')) {
        return { success: true, ignored: true, message: '管理中心不計入訪客觸擊' };
      }

      const now = new Date();

      const range = sheet.getDataRange();
      const values = range.getValues();
      let targetRow = -1;
      let currentViews = 0;
      let existingRowTitle = '';

      for (let r = 1; r < values.length; r++) {
        const existingPath = String(values[r][1] || '').trim();
        const existingPid = String(values[r][2] || '').trim();
        
        if (productId !== '-' && productId !== '') {
          if (existingPid === productId || existingPath === pagePath) {
            targetRow = r + 1;
            existingRowTitle = String(values[r][0] || '').trim();
            currentViews = Number(values[r][3]) || 0;
            break;
          }
        } else {
          if (existingPath === pagePath) {
            targetRow = r + 1;
            existingRowTitle = String(values[r][0] || '').trim();
            currentViews = Number(values[r][3]) || 0;
            break;
          }
        }
      }

      const newViews = currentViews + 1;

      if (targetRow > 0) {
        // 🛡️ 關鍵防護 2：嚴禁以「連線代購即時訂購單」覆蓋現有的真實商品名稱！
        const isNewGeneric = pageTitle.includes('連線代購即時訂購單') || pageTitle === '商品內頁';
        const isExistingGood = existingRowTitle && !existingRowTitle.includes('連線代購即時訂購單') && existingRowTitle !== '商品內頁';

        if (pageTitle && (!isExistingGood || !isNewGeneric)) {
          sheet.getRange(targetRow, 1).setValue(pageTitle);
        }
        if (productId && productId !== '-') sheet.getRange(targetRow, 3).setValue(productId);
        sheet.getRange(targetRow, 4).setValue(newViews);
        sheet.getRange(targetRow, 5).setValue(now);
      } else {
        sheet.appendRow([pageTitle, pagePath, productId, 1, now]);
      }

      // 📧 即時發送【網址觸擊 / 進站通報】Email 通知給管理者 (wen.chen2407@gmail.com)
      try {
        sendPageViewAlertEmail({
          pagePath: pagePath,
          pageTitle: pageTitle,
          productId: productId,
          newViews: newViews,
          userName: data.userName || '',
          userId: data.userId || '',
          device: data.device || '',
          now: now
        });
      } catch (mailErr) {
        console.warn('發送網址觸擊通知信略過:', mailErr);
      }

      return { success: true, page: pagePath, views: newViews };
    } catch (err) {
      return { success: false, message: err.toString() };
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

    // 先行載入商品清單以利標題自動校正與資訊豐富化 (圖片、價格)
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

    const currentValues = sheet.getDataRange().getValues();

    let totalViews = 0;
    const pages = [];
    const productViewsMap = {};

    for (let r = 1; r < currentValues.length; r++) {
      let title = String(currentValues[r][0] || '').trim();
      const path = String(currentValues[r][1] || '').trim();
      const pid = String(currentValues[r][2] || '').trim();
      const views = Number(currentValues[r][3]) || 0;
      const lastTime = currentValues[r][4] ? Utilities.formatDate(new Date(currentValues[r][4]), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : '';

      if (isExcludedPageView(path, title)) continue;

      // 🌟 自動校正並回寫試算表：若商品標題曾被暫存覆蓋成「連線代購即時訂購單」或「商品內頁」，自動換回商品真實名稱！
      if (pid && pid !== '-') {
        if (title.includes('連線代購即時訂購單') || title === '商品內頁' || !title) {
          if (productsMeta[pid] && productsMeta[pid].name) {
            title = `【商品】${productsMeta[pid].name}`;
            try {
              sheet.getRange(r + 1, 1).setValue(title);
            } catch (healErr) {
              console.warn('自動校正試算表標題失敗:', healErr);
            }
          }
        }
      }

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

/**
 * 試算表開啟時自動建立專屬操作選單 (支援試算表內一鍵點擊授權與測試、手動清理日誌)
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🤖 系統管理與 Telegram')
      .addItem('🚀 授權並測試 Telegram 連線', 'forceAuthAndTest')
      .addSeparator()
      .addItem('🧹 手動清理過期會員日誌 (保留最新5000筆)', 'cleanOldCustomerLogs')
      .addToUi();
  } catch (e) {}
}

/**
  * 🌟 試算表即時編輯觸發器 (Simple Trigger)：
  * 當賣家在【顧客歸戶】手動修改 C 欄「真實姓名」或 D 欄「電話」時（例如修正被舊版覆蓋的會員真實姓名），
  * 自動即時將最新的會員真實姓名與電話同步至【訂單明細】該會員所有訂單的第 23 欄 (W: 真實姓名) 與第 24 欄 (X: 會員電話)，
  * 且完全保留第 13/14 欄 (M/N: 結帳收件人姓名與電話) 不受影響！
  */
  function onEdit(e) {
    try {
      if (!e || !e.range) return;
      var sheet = e.range.getSheet();
      if (!sheet || sheet.getName() !== SHEET_NAMES.CUSTOMERS) return;

      var row = e.range.getRow();
      var col = e.range.getColumn();
      var numCols = e.range.getNumColumns();
      if (row <= 1) return;

      // 檢查是否編輯了第 3 欄 (C: 真實姓名) 或 第 4 欄 (D: 電話)
      var touchesNameOrPhone = (col <= 4 && (col + numCols - 1) >= 3);
      if (!touchesNameOrPhone) return;

      var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
      var orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
      if (!orderSheet) return;

      ensureOrderAndCustomerSchema(ss, orderSheet, sheet);

      var numRows = e.range.getNumRows();
      var custRows = sheet.getRange(row, 1, numRows, 4).getValues();
      var updatesByUid = {};

      for (var r = 0; r < custRows.length; r++) {
        var uid = String(custRows[r][0] || '').trim();
        var realName = String(custRows[r][2] || '').trim();
        var phone = String(custRows[r][3] || '').trim();
        if (uid) {
          updatesByUid[uid] = { realName: realName, phone: phone };
        }
      }

      var orderLastRow = orderSheet.getLastRow();
      if (orderLastRow <= 1) return;

      var orderData = orderSheet.getRange(2, 1, orderLastRow - 1, 24).getValues();
      for (var i = 0; i < orderData.length; i++) {
        var oUid = String(orderData[i][2] || '').trim();
        if (oUid && updatesByUid[oUid]) {
          var targetRow = i + 2;
          var u = updatesByUid[oUid];
          if (u.realName) {
            orderSheet.getRange(targetRow, 23).setValue(u.realName); // W 欄: 訂購會員真實姓名
          }
          if (u.phone) {
            orderSheet.getRange(targetRow, 24).setValue(formatPhoneAsText(u.phone)); // X 欄: 訂購會員電話
          }
        }
      }
    } catch (err) {
      console.warn('onEdit 同步會員真實姓名略過:', err);
    }
  }

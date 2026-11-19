/**
 * ==============================================================================
 * 【00_Config.gs】系統常數、白名單與共用基礎函式
 * ==============================================================================
 * 說明：存放全域工作表名稱、Google Drive 資料夾 ID、管理員白名單及基礎共用函式。
 * ==============================================================================
 */

/**
  * LINE 連線代購 & 群組團購整合系統 - GAS 後端 API
  * 專為 Google Sheets + LINE LIFF 架構設計
  */

  // 工作表名稱常數
  const SHEET_NAMES = {
    PRODUCTS: '商品清單',
    ORDERS: '訂單明細',
    CUSTOMERS: '顧客歸戶',
    CUSTOMER_LOGS: '會員動態Log',
    SETTINGS: '系統設定',
    VIEWS: '瀏覽統計'
  };

  // 指定的商品圖片 Google Drive 資料夾 ID
  const DRIVE_FOLDER_ID = '1dqoqS4VJK7Fn56-L-mM5Afwyb6dvTQMB';

  // 📧 管理者接收【網址觸擊 / 訪客進站】即時通報的 Email 信箱
  const ADMIN_ALERT_EMAIL = 'wen.chen2407@gmail.com';

  // 🛡️ 管理者 LINE User ID 白名單（最高權限白名單：可查看進貨成本、毛利計算機與顧客通知）
  const DEFAULT_ADMIN_LINE_UIDS = [
    'Uee7765ed3c1819cb94fd404cfab8553b' // 賣家主要管理員
  ];

  // 預設管理員安全通行碼 (忘記時可直接至 Google 試算表【系統設定】ADMIN_API_KEY 欄位查看或自訂修改)
  const DEFAULT_ADMIN_API_KEY = 'admin8888';

/**
  * 檢查傳入之 LINE User ID 與安全通行碼 (adminToken) 是否具備管理者權限
  * 雙重防護機制：
  * 1. LINE User ID 必須符合常數白名單或試算表【系統設定】中的 ADMIN_LINE_UIDS。
  * 2. 同時必須傳送正確的安全通行碼 (adminToken)，與試算表【系統設定】中的 ADMIN_API_KEY 相符。
  * 徹底杜絕單憑明文 LINE UID 就能偽造請求取得個資之重大資安漏洞！
  */
  function checkIsAdmin(userId, adminToken) {
    if (!userId) return false;
    const uid = String(userId).trim();
    const token = String(adminToken || '').trim();

    let isUidWhitelisted = DEFAULT_ADMIN_LINE_UIDS.includes(uid);
    let expectedToken = DEFAULT_ADMIN_API_KEY;

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const settSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
      if (settSheet) {
        const data = settSheet.getDataRange().getValues();
        for (let r = 1; r < data.length; r++) {
          const key = String(data[r][0] || '').trim();
          const val = String(data[r][1] || '').trim();
          if (key === 'ADMIN_LINE_UIDS') {
            const list = val.split(/[,\n;|\s]+/).map(s => s.trim()).filter(Boolean);
            if (list.includes(uid)) isUidWhitelisted = true;
          } else if (key === 'ADMIN_API_KEY') {
            if (val) expectedToken = val;
          }
        }
      }
    } catch (err) {
      console.warn('檢查試算表管理者白名單或安全通行碼失敗:', err);
    }

    // 1. UID 必須在白名單內
    if (!isUidWhitelisted) return false;

    // 2. 雙重驗證：安全通行碼 (Token) 必須正確
    if (token && token === expectedToken) {
      return true;
    }

    return false;
  }

/**
  * 工具函式：輸出 JSON
  */
  function jsonResponse(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
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

/**
 * ==============================================================================
 * 【01_Router.gs】Web App HTTP 請求路由器 (doGet / doPost)
 * ==============================================================================
 * 說明：負責統一接收前端 fetch 請求並轉發給相應的 Service 模組處理。
 * ==============================================================================
 */

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
    const reqUserId = (e && e.parameter && (e.parameter.userId || e.parameter.uid || e.parameter.adminUserId)) || '';
    const reqAdminToken = (e && e.parameter && (e.parameter.adminToken || e.parameter.token || e.parameter.apiKey)) || '';
    
    try {
      let result = {};
      if (action === 'getProducts') {
        // 買家前台商品大廳公開 API：僅回傳上架中商品，嚴格排除採購成本、匯率與毛利等商業機密
        result = getProductsList(false, false);
      } else if (action === 'getAllProducts') {
        // 賣家後台管理中心專用：強制驗證管理者 LINE User ID 白名單與安全通行碼！
        const isAdmin = checkIsAdmin(reqUserId, reqAdminToken);
        if (!isAdmin) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = getProductsList(true, true);
          result.isAdmin = true;
        }
      } else if (action === 'verifyAdmin') {
        // 提供管理中心前端門禁系統快速校驗當前 LINE 使用者身分與安全通行碼
        const isAdmin = checkIsAdmin(reqUserId, reqAdminToken);
        result = { success: isAdmin, isAdmin: isAdmin, userId: reqUserId };
      } else if (action === 'getProductById') {
        const pid = e.parameter.productId;
        // 買家單品詳情頁公開 API：嚴格排除採購成本、匯率與毛利等商業機密
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
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = getViewStats();
        }
      } else if (action === 'testTelegram') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          const testRes = sendTelegramMessage('🔔 <b>【連線系統測試】</b>\n恭喜！您的 Telegram 機器人已成功與 Google Apps Script 建立即時連線！\n• <b>訪客稱呼</b>：系統管理員\n• <b>連線時間</b>：' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'));
          result = testRes;
        }
      } else if (action === 'recordPageView') {
        result = handleRecordPageView({
          page: e.parameter.page,
          title: e.parameter.title,
          productId: e.parameter.productId
        });
      } else if (action === 'getCustomerLogs') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = getCustomerLogs(e.parameter.limit || 100);
        }
      } else if (action === 'getCustomerNotifyList') {
        // 顧客結帳通知清單包含買家個資與消費，強制驗證管理者身分與安全通行碼
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = getCustomerNotifyList();
        }
      } else if (action === 'getCheckoutParcels') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = getCheckoutParcels();
        }
      } else if (action === 'getProcurementOrders') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = getProcurementOrders();
        }
      } else if (action === 'updateParcelStatus') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = handleUpdateParcelStatus({
            checkoutId: e.parameter.checkoutId,
            status: e.parameter.status,
            trackingCode: e.parameter.trackingCode
          });
        }
      } else if (action === 'recalculateCustomerStats') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = recalculateCustomerStats();
        }
      } else if (action === 'upgradeSheets' || action === 'upgradeOrderAndCustomerSheets') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = upgradeOrderAndCustomerSheets();
        }
      } else if (action === 'repairProductCostData' || action === 'upgradeCostData') {
        if (!checkIsAdmin(reqUserId, reqAdminToken)) {
          result = { success: false, message: '🔒 存取遭拒：管理者權限或安全通行碼驗證失敗！', isAdmin: false };
        } else {
          result = { success: true, message: upgradeExistingProductsToFormula() };
        }
      } else if (action === 'clearTestData' || action === 'clearTestProducts') {
        result = { success: false, message: '🚫 系統已啟用防誤觸保護：為防止訂單與顧客資料遭誤刪，此 API 已永久停用！' };
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
  * - addProduct (賣家快速上架商品 - 需管理員白名單權限與安全通行碼)
  * - updateProduct (賣家線上完整維護商品資料 - 需管理員白名單權限與安全通行碼)
  * - batchUpdateProducts (批次上架/下架/修改分類 - 需管理員白名單權限與安全通行碼)
  * - uploadImages (圖片上傳 Google Drive - 需管理員白名單權限與安全通行碼)
  * - markCustomerNotified (標記顧客名下訂單已發送1對1結帳通知 - 需管理員白名單權限與安全通行碼)
  */
  function doPost(e) {
    try {
      const postData = JSON.parse(e.postData.contents);
      const action = postData.action;
      const reqAdminId = (postData && (
        postData.adminUserId ||
        postData.userId ||
        (postData.data && (postData.data.adminUserId || postData.data.userId))
      )) || '';
      const reqAdminToken = (postData && (
        postData.adminToken ||
        postData.token ||
        postData.apiKey ||
        (postData.data && (postData.data.adminToken || postData.data.token || postData.data.apiKey))
      )) || '';

      // 🛡️ 後台管理專屬寫入與統計操作：強制驗證管理員 LINE UID 白名單與安全通行碼
      const adminProtectedActions = [
        'addProduct',
        'updateProduct',
        'batchUpdateProducts',
        'uploadImages',
        'markCustomerNotified',
        'updateParcelStatus',
        'getViewStats',
        'getCheckoutParcels',
        'getProcurementOrders'
      ];
      if (adminProtectedActions.includes(action) && !checkIsAdmin(reqAdminId, reqAdminToken)) {
        return jsonResponse({
          success: false,
          isAdmin: false,
          message: '🚫 權限不足：管理員身分或安全通行碼 (PIN) 驗證失敗，系統已拒絕執行此後台操作！'
        });
      }

      let result = {};
      if (action === 'createOrder') {
        result = handleCreateOrder(postData.data);
      } else if (action === 'registerCustomer') {
        result = handleRegisterCustomer(postData.data);
      } else if (action === 'checkoutOrders') {
        result = handleCheckoutOrders(postData.data);
      } else if (action === 'confirmReceived') {
        result = handleConfirmReceived(postData.data);
      } else if (action === 'updateParcelStatus') {
        result = handleUpdateParcelStatus(postData.data || postData);
      } else if (action === 'getCheckoutParcels') {
        result = getCheckoutParcels();
      } else if (action === 'getProcurementOrders') {
        result = getProcurementOrders();
      } else if (action === 'reportPayment') {
        result = handleReportPayment(postData.data);
      } else if (action === 'addProduct') {
        result = handleAddProduct(postData.data);
      } else if (action === 'updateProduct') {
        result = handleUpdateProduct(postData.data);
      } else if (action === 'batchUpdateProducts') {
        result = handleBatchUpdateProducts(postData.data);
      } else if (action === 'uploadImages') {
        result = handleUploadImages(postData.data);
      } else if (action === 'markCustomerNotified') {
        result = handleMarkCustomerNotified(postData);
      } else if (action === 'recordPageView') {
        result = handleRecordPageView(postData.data);
      } else if (action === 'getViewStats') {
        result = getViewStats();
      } else if (action === 'getCustomerLogs') {
        result = getCustomerLogs(postData.limit || 100);
      } else if (action === 'upgradeSheets' || action === 'upgradeOrderAndCustomerSheets') {
        result = { success: true, message: upgradeOrderAndCustomerSheets() };
      } else if (action === 'clearTestData' || action === 'clearTestProducts') {
        result = { success: false, message: '🚫 系統已啟用防誤觸保護：為防止訂單與顧客資料遭誤刪，此 API 已永久停用！' };
      } else {
        result = { success: false, message: '未知的 action' };
      }
      return jsonResponse(result);
    } catch (err) {
      return jsonResponse({ success: false, message: err.toString(), error: err.toString() });
    }
  }

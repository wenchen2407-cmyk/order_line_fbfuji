/**
 * ==============================================================================
 * 【02_ProductService.gs】商品管理服務模組 (CRUD & 成本計算)
 * ==============================================================================
 * 說明：專門處理商品列表取得、詳情、新增、編輯、上下架、多圖上傳與原幣成本計算。
 * ==============================================================================
 */

/**
  * 取得商品清單
  * @param {boolean} includeAll - 是否包含已下架商品 (預設 false: 僅回傳上架中商品)
  * @param {boolean} includeCost - 是否包含採購成本、匯率與利潤欄位 (預設 false: 嚴格禁止回傳商業機密)
  */
  function getProductsList(includeAll = false, includeCost = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    if (!sheet) return { success: false, message: '尚未初始化商品表' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const products = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[9];
      if (includeAll || status === '上架中') {
        let rawSpecs = [];
        try {
          rawSpecs = JSON.parse(row[6]);
        } catch (e) {
          rawSpecs = row[6] ? [row[6]] : [];
        }

        const basePrice = Number(row[4]) || 0;
        const baseStock = Number(row[5]) || 0;
        const isBaseUnlimited = baseStock >= 9999;

        // 正規化規格陣列 (相容舊版純字串陣列 與 新版獨立定價庫存物件陣列)
        let normalizedSpecs = [];
        if (Array.isArray(rawSpecs)) {
          normalizedSpecs = rawSpecs.map(s => {
            if (s && typeof s === 'object') {
              const sName = String(s.name || '').trim();
              const sPrice = (s.price !== undefined && !isNaN(Number(s.price)) && Number(s.price) > 0) ? Number(s.price) : basePrice;
              const sUnlimited = Boolean(s.isUnlimited || Number(s.stock) >= 9999);
              const sStock = sUnlimited ? 999999 : (Number(s.stock) || 0);
              return { name: sName, price: sPrice, stock: sStock, isUnlimited: sUnlimited };
            } else if (typeof s === 'string' && s.trim()) {
              return { name: s.trim(), price: basePrice, stock: baseStock, isUnlimited: isBaseUnlimited };
            }
            return null;
          }).filter(s => s && s.name);
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

        const createdAtStr = row[10] ? formatDeadlineStr(row[10]) : '';
        const updatedAtStr = row[21] ? formatDeadlineStr(row[21]) : (createdAtStr || '');

        const prodItem = {
          id: row[0],
          name: row[1],
          category: row[2],
          originalPrice: hasOrigPrice ? Number(origPriceVal) : 0,
          price: basePrice,
          stock: baseStock,
          specs: normalizedSpecs,
          imageUrl: row[7] || (allImages[0] || ''),
          imageUrls: allImages,
          description: row[8],
          status: status || '上架中',
          deadline: row[14] ? formatDeadlineStr(row[14]) : '', // 欄位 15: 各商品指定結單時間 (選填)
          createdAt: createdAtStr,                             // 欄位 11 (K): 建立時間
          updatedAt: updatedAtStr                              // 欄位 22 (V): 最後更新時間
        };

        // 🛡️ 商業機密嚴格保護：採購成本、匯率、稅金、重量與運費僅在賣家管理後台明確請求 (includeCost=true) 時才附加
        // 前台買家公開 API (getProducts, getProductById) 絕對不包含任何成本欄位，防止客人窺探定價與利潤！
        if (includeCost) {
          prodItem.costPrice = Number(row[11]) || 0;
          prodItem.costNote = row[12] || '';
          prodItem.currency = row[15] || 'JPY';
          prodItem.costPreTax = Number(row[16]) || 0;
          prodItem.costTax = Number(row[17]) || 0;
          prodItem.exchangeRate = Number(row[18]) || 0;
          prodItem.weightG = Number(row[19]) || 0;
          prodItem.shippingRate = Number(row[20]) || 240;
        }

        products.push(prodItem);
      }
    }

    const settings = getSystemSettings().data || {};
    return { success: true, data: products, settings: settings };
  }

/**
  * 格式化結單時間字串 (相容 Date 物件與字串)
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
  * 取得單一商品詳情 (前台買家專用：嚴格禁止回傳任何採購成本與利潤資訊)
  */
  function getProductDetail(productId) {
    const list = getProductsList(false, false);
    if (!list.success) return list;
    const prod = list.data.find(p => p.id === productId);
    if (prod) {
      return { success: true, data: prod };
    }
    return { success: false, message: '查無此商品或已下架' };
  }

/**
  * 自動產生商品編號：
  * 日本代購: J + 西元年後2碼 + 日期(MMdd) + 0001 起流水號 (如 J2609270001)
  * 美國代購: U + 西元年後2碼 + 日期(MMdd) + 0001 起流水號 (如 U2610210001)
  * 韓國代購: K + 西元年後2碼 + 日期(MMdd) + 0001 起流水號 (如 K2609270001)
  * 台灣現貨: T + 西元年後2碼 + 日期(MMdd) + 0001 起流水號 (如 T2609270001)
  */
  function generateProductId(sheet, currency, now) {
    const cur = String(currency || 'JPY').toUpperCase();
    let prefixChar = 'J';
    if (cur === 'USD' || cur.includes('美') || cur.includes('USA')) {
      prefixChar = 'U';
    } else if (cur === 'KRW' || cur.includes('韓') || cur.includes('KOR')) {
      prefixChar = 'K';
    } else if (cur === 'TWD' || cur.includes('台') || cur.includes('TW')) {
      prefixChar = 'T';
    } else {
      prefixChar = 'J';
    }
    const datePart = Utilities.formatDate(now || new Date(), 'Asia/Taipei', 'yyMMdd');
    const basePrefix = prefixChar + datePart; // 例如 J260927, U261021, K260927, T260927

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
        if (hint.includes('USD') || hint.includes('美') || hint.includes('USA')) {
          cur = 'USD';
        } else if (hint.includes('KRW') || hint.includes('韓') || hint.includes('KOR')) {
          cur = 'KRW';
        } else if (hint.includes('TWD') || hint.includes('台') || hint.includes('TW')) {
          cur = 'TWD';
        } else {
          cur = 'JPY';
        }
      }

      // 商品編號：若賣家有自訂則使用自訂，否則依「J/U/K/T + 西元年後2碼 + 日期 + 0001」自動跳號
      const pid = (data.id && String(data.id).trim()) || generateProductId(sheet, cur, now);

      // 規格處理 (支援獨立定價/庫存物件陣列、純字串陣列或逗號分隔字串)
      let specsArray = [];
      const basePrice = Number(data.price) || 0;
      let calculatedStock = 0;
      let hasExplicitSpecs = false;

      if (Array.isArray(data.specs)) {
        specsArray = data.specs.map(item => {
          if (item && typeof item === 'object') {
            hasExplicitSpecs = true;
            const sName = String(item.name || '').trim();
            const sPrice = (item.price !== undefined && !isNaN(Number(item.price)) && Number(item.price) > 0) ? Number(item.price) : basePrice;
            const sUnlimited = Boolean(item.isUnlimited || Number(item.stock) >= 9999);
            const sStock = sUnlimited ? 999999 : (Number(item.stock) || 0);
            if (!sUnlimited) calculatedStock += sStock;
            return { name: sName, price: sPrice, stock: sStock, isUnlimited: sUnlimited };
          } else if (typeof item === 'string' && item.trim()) {
            return { name: item.trim(), price: basePrice, stock: data.isUnlimitedStock ? 999999 : (Number(data.stock) || 10), isUnlimited: Boolean(data.isUnlimitedStock) };
          }
          return null;
        }).filter(s => s && s.name);
      } else if (typeof data.specs === 'string' && data.specs.trim()) {
        specsArray = data.specs.split(/[,，\n]/).map(s => s.trim()).filter(Boolean).map(name => ({
          name: name,
          price: basePrice,
          stock: data.isUnlimitedStock ? 999999 : (Number(data.stock) || 10),
          isUnlimited: Boolean(data.isUnlimitedStock)
        }));
      }

      // 圖片處理：支援多張圖片（第一張為主圖）
      let imagesArray = [];
      if (Array.isArray(data.imageUrls) && data.imageUrls.length > 0) {
        imagesArray = data.imageUrls;
      } else if (data.imageUrl) {
        imagesArray = data.imageUrl.split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      }
      const mainImage = imagesArray[0] || data.imageUrl || '';

      // 庫存名額：若有獨立規格庫存，且有任一規格為不限庫存則總庫存為 999999；否則為各規格加總
      let stockQty = data.isUnlimitedStock ? 999999 : (Number(data.stock) || 10);
      if (hasExplicitSpecs && specsArray.length > 0) {
        const anyUnlimited = specsArray.some(s => s.isUnlimited);
        stockQty = anyUnlimited ? 999999 : calculatedStock;
      }

      const targetRow = sheet.getLastRow() + 1;

      // 採購原幣、稅金、匯率與重量獨立分開登記 (嚴格以實際輸入資料為準，未輸入則保留空白)
      if (data.currency) cur = String(data.currency).toUpperCase();
      const hasPreTax = data.costPricePreTax !== undefined && data.costPricePreTax !== '' && !isNaN(Number(data.costPricePreTax));
      const hasTax = data.taxAmount !== undefined && data.taxAmount !== '' && !isNaN(Number(data.taxAmount));
      const hasRate = data.exchangeRate !== undefined && data.exchangeRate !== '' && !isNaN(Number(data.exchangeRate));
      const hasWeight = data.weightG !== undefined && data.weightG !== '' && !isNaN(Number(data.weightG));

      const preTax = hasPreTax ? Number(data.costPricePreTax) : '';
      const tax = hasTax ? Number(data.taxAmount) : '';
      const rate = hasRate ? Number(data.exchangeRate) : '';
      const weight = hasWeight ? Number(data.weightG) : '';
      const shipRate = (data.shippingRate !== undefined && data.shippingRate !== '') ? Number(data.shippingRate) : 240;

      // 成本價 (NT$)：精準分開四捨五入公式 ROUND((Q+R)*S) + ROUND((T/1000)*U)
      let costCellVal = Number(data.costPrice) || 0;
      if (Number(preTax) > 0 || Number(tax) > 0 || Number(weight) > 0) {
        costCellVal = `=IF(OR(N(Q${targetRow})>0, N(T${targetRow})>0), ROUND((N(Q${targetRow}) + N(R${targetRow})) * N(S${targetRow})) + ROUND((N(T${targetRow}) / 1000) * IF(N(U${targetRow})>0, N(U${targetRow}), 240)), ${costCellVal})`;
      }

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
        costCellVal,                 // 欄位 12 (L): 成本價 (NT$) -> 支援試算表公式連動
        data.costNote || '',         // 欄位 13 (M): 採購備註說明
        JSON.stringify(imagesArray), // 欄位 14 (N): 所有圖片清單
        data.deadline ? String(data.deadline).trim() : '', // 欄位 15 (O): 指定結單時間
        cur,                         // 欄位 16 (P): 採購幣別
        preTax,                      // 欄位 17 (Q): 原幣稅前金額
        tax,                         // 欄位 18 (R): 稅金 (原幣)
        rate,                        // 欄位 19 (S): 採購匯率
        weight,                      // 欄位 20 (T): 商品重量 (g)
        shipRate,                    // 欄位 21 (U): 國際運費單價 (元/kg)
        Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') // 欄位 22 (V): 最後更新時間
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
  * 賣家線上完整維護商品資料 (100% 全欄位對齊建檔表單，並更新最後更新時間)
  */
  function handleUpdateProduct(data) {
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

      const pid = String(data.id || '').trim();
      if (!pid) return { success: false, message: '缺少商品編號 (ID)' };

      if (!data.name || data.price === undefined || data.price === '') {
        return { success: false, message: '商品名稱與價格為必填欄位！' };
      }

      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return { success: false, message: '商品清單為空' };

      // 尋找商品所在的列 (row)
      const idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let rowIndex = -1;
      for (let i = 0; i < idRange.length; i++) {
        if (String(idRange[i][0]).trim() === pid) {
          rowIndex = i + 2; // 試算表從第 2 列開始 (1-based)
          break;
        }
      }

      if (rowIndex === -1) {
        return { success: false, message: `找不到商品編號：${pid}` };
      }

      const nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

      // 1. 規格處理 (支援獨立定價/庫存物件陣列、純字串陣列或逗號分隔字串)
      let specsArray = [];
      const basePrice = Number(data.price) || 0;
      let calculatedStock = 0;
      let hasExplicitSpecs = false;

      if (Array.isArray(data.specs)) {
        specsArray = data.specs.map(item => {
          if (item && typeof item === 'object') {
            hasExplicitSpecs = true;
            const sName = String(item.name || '').trim();
            const sPrice = (item.price !== undefined && !isNaN(Number(item.price)) && Number(item.price) > 0) ? Number(item.price) : basePrice;
            const sUnlimited = Boolean(item.isUnlimited || Number(item.stock) >= 9999);
            const sStock = sUnlimited ? 999999 : (Number(item.stock) || 0);
            if (!sUnlimited) calculatedStock += sStock;
            return { name: sName, price: sPrice, stock: sStock, isUnlimited: sUnlimited };
          } else if (typeof item === 'string' && item.trim()) {
            return { name: item.trim(), price: basePrice, stock: data.isUnlimitedStock ? 999999 : (Number(data.stock) || 10), isUnlimited: Boolean(data.isUnlimitedStock) };
          }
          return null;
        }).filter(s => s && s.name);
      } else if (typeof data.specs === 'string' && data.specs.trim()) {
        specsArray = data.specs.split(/[,，\n]/).map(s => s.trim()).filter(Boolean).map(name => ({
          name: name,
          price: basePrice,
          stock: data.isUnlimitedStock ? 999999 : (Number(data.stock) || 10),
          isUnlimited: Boolean(data.isUnlimitedStock)
        }));
      }

      // 2. 圖片處理：支援多張圖片（第一張為主圖）
      let imagesArray = [];
      if (Array.isArray(data.imageUrls) && data.imageUrls.length > 0) {
        imagesArray = data.imageUrls.filter(Boolean);
      } else if (data.imageUrl) {
        imagesArray = String(data.imageUrl).split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      }
      const mainImage = imagesArray[0] || data.imageUrl || '';

      // 3. 庫存名額計算
      let stockQty = data.isUnlimitedStock ? 999999 : (Number(data.stock) || 0);
      if (hasExplicitSpecs && specsArray.length > 0) {
        const anyUnlimited = specsArray.some(s => s.isUnlimited);
        stockQty = anyUnlimited ? 999999 : calculatedStock;
      }

      // 4. 成本與公式處理
      let cur = (data.currency || 'JPY').toUpperCase();
      const hasPreTax = data.costPricePreTax !== undefined && data.costPricePreTax !== '' && !isNaN(Number(data.costPricePreTax));
      const hasTax = data.taxAmount !== undefined && data.taxAmount !== '' && !isNaN(Number(data.taxAmount));
      const hasRate = data.exchangeRate !== undefined && data.exchangeRate !== '' && !isNaN(Number(data.exchangeRate));
      const hasWeight = data.weightG !== undefined && data.weightG !== '' && !isNaN(Number(data.weightG));

      const preTax = hasPreTax ? Number(data.costPricePreTax) : '';
      const tax = hasTax ? Number(data.taxAmount) : '';
      const rate = hasRate ? Number(data.exchangeRate) : '';
      const weight = hasWeight ? Number(data.weightG) : '';
      const shipRate = (data.shippingRate !== undefined && data.shippingRate !== '') ? Number(data.shippingRate) : 240;

      let costCellVal = Number(data.costPrice) || 0;
      let costFormula = '';
      if (Number(preTax) > 0 || Number(tax) > 0 || Number(weight) > 0) {
        costFormula = `=IF(OR(N(Q${rowIndex})>0, N(T${rowIndex})>0), ROUND((N(Q${rowIndex}) + N(R${rowIndex})) * N(S${rowIndex})) + ROUND((N(T${rowIndex}) / 1000) * IF(N(U${rowIndex})>0, N(U${rowIndex}), 240)), ${costCellVal})`;
      }

      const statusVal = data.status === '已下架' ? '已下架' : '上架中';

      // 確保欄位至少有 22 欄
      if (sheet.getMaxColumns() < 22) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), 22 - sheet.getMaxColumns());
      }

      // 更新各欄位 (嚴格對齊 22 欄規格)
      sheet.getRange(rowIndex, 2).setValue(String(data.name || '').trim()); // B: 商品名稱
      sheet.getRange(rowIndex, 3).setValue(String(data.category || '連線好物').trim()); // C: 分類
      sheet.getRange(rowIndex, 4).setValue((data.originalPrice && Number(data.originalPrice) > 0) ? Number(data.originalPrice) : ''); // D: 專櫃原價
      sheet.getRange(rowIndex, 5).setValue(basePrice); // E: 連線代購價
      sheet.getRange(rowIndex, 6).setValue(stockQty); // F: 現貨庫存
      sheet.getRange(rowIndex, 7).setValue(JSON.stringify(specsArray)); // G: 規格清單(JSON)
      sheet.getRange(rowIndex, 8).setValue(mainImage); // H: 封面主圖網址
      sheet.getRange(rowIndex, 9).setValue(data.description || ''); // I: 商品描述
      sheet.getRange(rowIndex, 10).setValue(statusVal); // J: 狀態
      // K 欄 (建立時間) 保留原值不改動
      if (costFormula) {
        sheet.getRange(rowIndex, 12).setFormula(costFormula); // L: 成本價公式
      } else {
        sheet.getRange(rowIndex, 12).setValue(costCellVal);
      }
      sheet.getRange(rowIndex, 13).setValue(data.costNote || ''); // M: 採購備註說明
      sheet.getRange(rowIndex, 14).setValue(JSON.stringify(imagesArray)); // N: 所有圖片清單
      sheet.getRange(rowIndex, 15).setValue(data.deadline ? String(data.deadline).trim() : ''); // O: 指定結單時間
      sheet.getRange(rowIndex, 16).setValue(cur); // P: 採購幣別
      sheet.getRange(rowIndex, 17).setValue(preTax); // Q: 原幣稅前金額
      sheet.getRange(rowIndex, 18).setValue(tax); // R: 稅金(原幣)
      sheet.getRange(rowIndex, 19).setValue(rate); // S: 採購匯率
      sheet.getRange(rowIndex, 20).setValue(weight); // T: 商品重量(g)
      sheet.getRange(rowIndex, 21).setValue(shipRate); // U: 國際運費單價
      sheet.getRange(rowIndex, 22).setValue(nowStr); // V: 最後更新時間

      SpreadsheetApp.flush();

      return {
        success: true,
        message: `商品 [${pid}] 更新成功！`,
        productId: pid,
        updatedAt: nowStr
      };
    } catch (err) {
      return { success: false, message: '更新商品失敗：' + err.toString() };
    } finally {
      lock.releaseLock();
    }
  }

/**
  * 批次更新商品 (批次上架、批次下架、批次修改分類)
  * data: {
  *   productIds: ['J2610...', 'J2610...'],
  *   updateType: 'status' | 'category',
  *   value: '上架中' | '已下架' | '新分類標籤'
  * }
  */
  function handleBatchUpdateProducts(data) {
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

      const pids = Array.isArray(data.productIds) ? data.productIds.map(id => String(id).trim()).filter(Boolean) : [];
      if (pids.length === 0) return { success: false, message: '請至少選擇一件商品！' };

      const updateType = data.updateType; // 'status' 或 'category'
      const value = String(data.value || '').trim();

      if (!updateType || !value) {
        return { success: false, message: '缺少更新類型或更新數值' };
      }

      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return { success: false, message: '商品清單為空' };

      if (sheet.getMaxColumns() < 22) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), 22 - sheet.getMaxColumns());
      }

      const nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      const idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

      let updatedCount = 0;
      const targetCol = updateType === 'status' ? 10 : (updateType === 'category' ? 3 : 0);
      if (targetCol === 0) return { success: false, message: '不支援的批次更新類型：' + updateType };

      for (let i = 0; i < idRange.length; i++) {
        const currentPid = String(idRange[i][0]).trim();
        if (pids.includes(currentPid)) {
          const row = i + 2;
          sheet.getRange(row, targetCol).setValue(value);
          sheet.getRange(row, 22).setValue(nowStr); // V欄: 最後更新時間
          updatedCount++;
        }
      }

      SpreadsheetApp.flush();

      return {
        success: true,
        message: `成功批次更新 ${updatedCount} 件商品！`,
        updatedCount: updatedCount,
        updateType: updateType,
        value: value,
        updatedAt: nowStr
      };
    } catch (err) {
      return { success: false, message: '批次更新失敗：' + err.toString() };
    } finally {
      lock.releaseLock();
    }
  }

/**
  * 處理附件圖片上傳至 Google Drive (極速優化版：利用資料夾繼承權限，省去逐檔遠端授權延遲)
  */
  function handleUploadImages(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
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
        const directUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
        uploadedUrls.push(directUrl);
      }

      return { success: true, urls: uploadedUrls };
    } catch (err) {
      return { success: false, message: '圖片上傳至 Google Drive 失敗：' + err.toString() };
    } finally {
      lock.releaseLock();
    }
  }

/**
  * 測試並激活 Google Drive 權限（可在 GAS 編輯器上方下拉選單直接點「執行」測試）
  */
  function testDrivePermission() {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      Logger.log('✅ 成功連接指定之 Google Drive 資料夾：' + folder.getName());
      return 'OK: ' + folder.getName();
    } catch (e) {
      const folderName = 'LINE代購_商品圖片庫';
      const folders = DriveApp.getFoldersByName(folderName);
      const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log('✅ 已成功授權 Google Drive，並自動連結/建立「' + folderName + '」資料夾！(ID: ' + folder.getId() + ')');
      return 'OK (Auto Folder): ' + folder.getName();
    }
  }

/**
  * 🌟 智慧修復與升級既有商品成本資料：
  * 1. 徹底解決 20 欄升級為 21 欄（新增稅金欄位）所造成的資料向左錯位問題（原本匯率 0.22 錯移到稅金、重量 200 錯移到匯率、運費 240 錯移到重量、運費單價為空）。
  * 2. 徹底消除 480102 天價異常錯誤（原公式將原幣 2400 乘上錯位重量 200）。
  * 3. 依據【採購備註說明】（例如：JPY 2400 (稅前2400+稅0), 重200g, 運費NT$48）權威還原：
  *    - 採購幣別 P (16): JPY
  *    - 原幣稅前金額 Q (17): 2400
  *    - 稅金(原幣) R (18): 0 (純數字 0)
  *    - 採購匯率 S (19): 0.22
  *    - 商品重量(g) T (20): 200
  *    - 國際運費(元/kg) U (21): 240
  * 4. 成本價 L (12) 全面套用精確分開四捨五入公式：
  *    =ROUND((Q + R) * S) + ROUND((T / 1000) * U)
  *    在該商品中計算結果保證精準為 576 元（528 + 48）。
  */
  function upgradeExistingProductsToFormula(targetSheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prodSheet = targetSheet || ss.getSheetByName(SHEET_NAMES.PRODUCTS);
    if (!prodSheet) return '找不到商品清單工作表';

    const lastRow = prodSheet.getLastRow();
    if (lastRow <= 1) return '目前尚無商品資料';

    // 確保表格寬度至少有 22 欄
    const maxCols = prodSheet.getMaxColumns();
    if (maxCols < 22) {
      prodSheet.insertColumnsAfter(maxCols, 22 - maxCols);
    }

    // 確保第 1 列表頭為最新 22 欄標準名稱
    const prodHeaders = [
      '商品編號', '商品名稱', '分類', '專櫃原價', '連線代購價', '現貨庫存', 
      '規格清單(JSON)', '封面主圖網址', '商品描述', '狀態', '建立時間',
      '成本價(NT$)', '採購備註說明', '所有圖片清單(JSON)', '指定結單時間',
      '採購幣別', '原幣稅前金額', '稅金(原幣)', '採購匯率', '商品重量(g)', '國際運費(元/kg)',
      '最後更新時間'
    ];
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setValues([prodHeaders]);
    prodSheet.getRange(1, 1, 1, prodHeaders.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');

    // 讀取第 2 列到最後一列所有商品資料 (22 欄)
    const range = prodSheet.getRange(2, 1, lastRow - 1, 22);
    const values = range.getValues();
    const formulas = range.getFormulas();

    let upgradedCount = 0;

    for (let i = 0; i < values.length; i++) {
      const rowNum = i + 2;
      const row = values[i];
      const pid = String(row[0] || '').trim();
      if (!pid) continue;

      const prodName = String(row[1] || '');
      const currentCost = row[11];
      const note = String(row[12] || '').trim();

      let cur = String(row[15] || '').trim().toUpperCase();
      let preTax = row[16];
      let tax = row[17];
      let rate = row[18];
      let weight = row[19];
      let shipRate = row[20];

      // 1. 偵測是否遭遇 20 欄擴展 21 欄造成的資料錯位 (Shifted)
      // 典型特徵：稅金 R 欄是小數匯率 (0.001~1.0)，而匯率 S 欄是大數字 (>=10，例如 200, 380)
      const isShifted = (typeof tax === 'number' && tax > 0 && tax <= 1.0) && 
                        (typeof rate === 'number' && rate >= 10);
      
      if (isShifted) {
        // 修正錯位：原 R 欄其實是匯率，原 S 欄其實是重量，原 T 欄其實是運費單價
        const shiftedRate = tax;
        const shiftedWeight = rate;
        const shiftedShip = (typeof weight === 'number' && weight > 0) ? weight : 240;

        rate = shiftedRate;
        weight = shiftedWeight;
        shipRate = shiftedShip;
        tax = 0;
      }

      // 2. 從「採購備註說明」(M 欄) 權威還原真實數據
      if (note) {
        // 解析幣別
        const curMatch = note.match(/(JPY|USD|KRW|TWD)/i);
        if (curMatch) {
          cur = curMatch[1].toUpperCase();
        } else if (!cur) {
          if (pid.startsWith('U') || prodName.includes('美') || note.includes('美') || note.includes('USD')) cur = 'USD';
          else if (pid.startsWith('K') || prodName.includes('韓') || note.includes('韓')) cur = 'KRW';
          else if (pid.startsWith('T') || prodName.includes('台') || note.includes('台')) cur = 'TWD';
          else cur = 'JPY';
        }

        // 解析稅前金額與稅金 (例：稅前2400+稅0 或 稅前3000+稅300)
        const splitMatch = note.match(/稅前\s*(\d+(?:\.\d+)?)\s*\+\s*稅(?:金)?\s*(\d+(?:\.\d+)?)/i);
        if (splitMatch) {
          preTax = Number(splitMatch[1]);
          tax = Number(splitMatch[2]);
        } else {
          // 單一金額 (例：JPY 2400, USD 120, 或 3500(含稅))
          const singleMatch = note.match(/(?:JPY|USD|KRW|TWD)?\s*(\d+(?:\.\d+)?)\s*(?:\(|含稅|稅前|\,|$)/i);
          if (singleMatch && (preTax === undefined || preTax === '' || isNaN(Number(preTax)) || Number(preTax) === 0 || isShifted)) {
            preTax = Number(singleMatch[1]);
            if (tax === undefined || tax === '' || isNaN(Number(tax)) || isShifted) {
              tax = 0;
            }
          }
        }

        // 解析商品重量 (例：重200g, 380g)
        const weightMatch = note.match(/(?:重|重量|重約|weight|[\,，\s])\s*(\d+(?:\.\d+)?)\s*g/i) || note.match(/(\d+)\s*g/i);
        if (weightMatch) {
          weight = Number(weightMatch[1]);
        }

        // 解析運費單價 (例：運費NT$48 且重量為 200g ➜ 48 / 0.2 = 240)
        if (note.includes('運費NT$') && weight > 0) {
          const freightMatch = note.match(/運費NT\$\s*(\d+)/i);
          if (freightMatch) {
            const freightVal = Number(freightMatch[1]);
            shipRate = Math.round(freightVal / (weight / 1000));
          }
        }
      }

      // 3. 確保預設值與格式正確
      if (!cur) cur = 'JPY';

      // 匯率合理性防護：若匯率不存在或殘留錯位重量，還原為幣別基準匯率
      if (cur === 'USD') {
        if (typeof rate !== 'number' || isNaN(rate) || rate <= 0 || rate < 10 || rate > 60) rate = 32.5;
      } else if (cur === 'KRW') {
        if (typeof rate !== 'number' || isNaN(rate) || rate <= 0 || rate >= 1) rate = 0.026;
      } else if (cur === 'TWD') {
        rate = 1.0;
      } else {
        // JPY
        if (typeof rate !== 'number' || isNaN(rate) || rate <= 0 || rate >= 1) rate = 0.22;
      }

      // 運費單價防護
      if (typeof shipRate !== 'number' || isNaN(shipRate) || shipRate <= 0) {
        shipRate = 240;
      }

      // 數值型別標準化
      const finalPreTax = (preTax !== undefined && preTax !== '' && !isNaN(Number(preTax))) ? Number(preTax) : '';
      const finalTax = (tax !== undefined && tax !== '' && !isNaN(Number(tax))) ? Number(tax) : (finalPreTax !== '' ? 0 : '');
      const finalWeight = (weight !== undefined && weight !== '' && !isNaN(Number(weight))) ? Number(weight) : '';

      // 寫回 values 矩陣
      row[15] = cur;
      row[16] = finalPreTax;
      row[17] = finalTax;
      row[18] = rate;
      row[19] = finalWeight;
      row[20] = shipRate;

      // 確保第 22 欄 (最後更新時間) 有初始值
      if (!row[21]) {
        const createdVal = row[10] || new Date();
        row[21] = createdVal instanceof Date ? Utilities.formatDate(createdVal, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : String(createdVal);
      }

      // 取得原本手動登記之真實成本數字 (防止 480102 等天價殘留為 baseCost)
      let baseCost = 0;
      if (typeof currentCost === 'number' && !isNaN(currentCost) && currentCost > 0 && currentCost < 50000) {
        baseCost = currentCost;
      }

      // 🌟 嚴格套用正確的成本價公式：
      // 商品折台幣 = ROUND((原幣稅前 Q + 稅金 R) * 匯率 S)
      // 國際運費台幣 = ROUND((重量 T / 1000) * 運費 U)
      // 兩者嚴格分開計算四捨五入！
      formulas[i][11] = `=IF(OR(N(Q${rowNum})>0, N(T${rowNum})>0), ROUND((N(Q${rowNum}) + N(R${rowNum})) * N(S${rowNum})) + ROUND((N(T${rowNum}) / 1000) * IF(N(U${rowNum})>0, N(U${rowNum}), 240)), ${baseCost})`;

      upgradedCount++;
    }

    // 批次寫入所有數值
    range.setValues(values);

    // 逐列覆寫 L 欄 (成本價) 公式
    for (let i = 0; i < formulas.length; i++) {
      if (formulas[i][11] && formulas[i][11].startsWith('=')) {
        prodSheet.getRange(i + 2, 12).setFormula(formulas[i][11]);
      }
    }

    SpreadsheetApp.flush();
    Logger.log(`[校正與修復完成] 共更新 ${upgradedCount} 筆商品！已徹底排除錯位，精準校正稅金/匯率/重量/運費並套用正確公式！`);
    return `成功校正並修復 ${upgradedCount} 筆商品！欄位錯位已排除，稅金與匯率已還原，成本價公式已全面校準。`;
  }

/**
  * 🌟 一鍵修復商品成本與欄位錯位（供 Apps Script 介面選取執行的別名函式）
  */
  function repairProductCostData() {
    return upgradeExistingProductsToFormula();
  }

/**
  * 🛠️ 清除自動化測試商品（已啟用防誤觸保護）
  */
  function clearTestProducts(safetyKey) {
    // 🛡️ 強制防誤觸檢查
    if (safetyKey !== 'CONFIRM_DELETE_PRODUCTS_2026') {
      Logger.log('🚫【安全防護已啟動】拒絕執行！為防止商品資料遭誤刪，已阻止清除動作。');
      return {
        success: false,
        message: '🚫【安全防護已啟動】拒絕執行！為保護商品庫，已阻止清除動作。'
      };
    }

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

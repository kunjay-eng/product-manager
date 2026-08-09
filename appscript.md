## Code.gs
/**
 * ============================================================
 * ระบบจัดการสินค้า — Code.gs ฉบับรวมทุกเฟส (Phase 2-10)
 * นี่คือไฟล์ Code.gs ไฟล์เดียวจบ — คัดลอกทั้งหมดนี้ไปแทนที่ Code.gs เดิมในโปรเจกต์
 * (ลบโค้ดเก่าทั้งหมดในไฟล์ Code.gs ออกก่อน แล้ววางไฟล์นี้แทน)
 *
 * ต้องมีไฟล์ HTML อีก 4 ไฟล์ในโปรเจกต์เดียวกัน (ชื่อต้องตรงเป๊ะ ตัวพิมพ์ใหญ่-เล็กสำคัญ):
 *   Index, Stylesheet, JavaScript, AccessDenied
 *
 * ทำตามลำดับนี้หลังวางโค้ด:
 *   1. รัน setupSheets() ครั้งเดียว (สร้าง 6 ชีตหลัก)
 *   2. รัน migratePurchaseStatusColumn() ครั้งเดียว (เพิ่มคอลัมน์ status ให้ Purchases)
 *   3. Deploy → New deployment → Web app
 *        Execute as: "User accessing the web app"
 *        Who has access: "Anyone with a Google account" (หรือ "Anyone within [องค์กร]" ถ้าใช้ Workspace)
 *   4. เปิดลิงก์เว็บแอป ครั้งแรกจะมีหน้าต่างขอ authorize สิทธิ์ (Sheets + Drive) กด Allow
 * ============================================================
 */


/* ============================================================
   PHASE 2 — โครงสร้าง Sheets + กันข้อมูลซ้ำ
   ============================================================ */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = {
    Categories: ['id', 'name', 'active'],
    ProductGroups: ['id', 'category_id', 'name', 'image_url', 'current_sell_price', 'active'],
    Stores: ['id', 'name', 'active'],
    Units: ['id', 'name', 'type'],
    Purchases: ['id', 'date', 'product_group_id', 'store_id', 'order_no', 'recorder', 'note',
                'qty_buy', 'unit_buy', 'ratio', 'unit_sell',
                'price_buy', 'shipping', 'discount', 'other_cost',
                'receipt_img', 'bill_img', 'created_at'],
    PriceHistory: ['id', 'product_group_id', 'date', 'cost_price', 'old_price', 'new_price',
                   'profit', 'margin', 'edited_by']
  };

  Object.keys(sheets).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    sheet.getRange(1, 1, 1, sheets[name].length).setValues([sheets[name]]);
    sheet.setFrozenRows(1);
  });

  const def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  SpreadsheetApp.flush();
  Logger.log('Setup done');
}

function generateId(sheetName, prefix) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + '001';

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const nums = ids
    .map(id => parseInt(String(id).replace(prefix, '')))
    .filter(n => !isNaN(n));

  const max = nums.length ? Math.max(...nums) : 0;
  return prefix + String(max + 1).padStart(3, '0');
}

function normalizeName(name) {
  return String(name).trim().toLowerCase().replace(/\s+/g, '');
}

function searchMaster(sheetName, keyword) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nameIdx = headers.indexOf('name');
  const activeIdx = headers.indexOf('active');

  const normKeyword = normalizeName(keyword);

  return data
    .filter(row => {
      if (activeIdx !== -1 && row[activeIdx] === false) return false;
      return normalizeName(row[nameIdx]).includes(normKeyword);
    })
    .map(row => ({ id: row[0], name: row[nameIdx] }));
}

function findExactMatch(sheetName, name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nameIdx = headers.indexOf('name');

  const normTarget = normalizeName(name);
  const found = data.find(row => normalizeName(row[nameIdx]) === normTarget);
  return found ? found[0] : null;
}

function findOrCreateCategory(name) {
  const existing = findExactMatch('Categories', name);
  if (existing) return existing;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const id = generateId('Categories', 'CAT');
  sheet.appendRow([id, name.trim(), true]);
  return id;
}

function findOrCreateProductGroup(categoryId, name, imageUrl, productUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const urlIdx = headers.indexOf('product_url');

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const normTarget = normalizeName(name);
    const foundIdx = data.findIndex(row => row[1] === categoryId && normalizeName(row[2]) === normTarget);
    if (foundIdx !== -1) {
      const found = data[foundIdx];
      if (productUrl && urlIdx !== -1 && !found[urlIdx]) {
        sheet.getRange(foundIdx + 2, urlIdx + 1).setValue(productUrl);
      }
      return found[0];
    }
  }

  const id = generateId('ProductGroups', 'PG');
  sheet.appendRow([id, categoryId, name.trim(), imageUrl || '', 0, true]);
  if (productUrl && urlIdx !== -1) {
    sheet.getRange(sheet.getLastRow(), urlIdx + 1).setValue(productUrl);
  }
  return id;
}


//**1) รันครั้งเดียว: เพิ่มคอลัมน์**
function migrateProductUrlColumn(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('product_url') !== -1){ Logger.log('มีคอลัมน์แล้ว'); return; }
  sheet.getRange(1, sheet.getLastColumn() + 1, 1, 1).setValues([['product_url']]);
  Logger.log('เพิ่มคอลัมน์ product_url แล้ว');
}


function migratePurchaseStatusColumn(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const newCols = ['status', 'cancel_reason', 'cancelled_at'].filter(c => headers.indexOf(c) === -1);
  if (newCols.length === 0){ Logger.log('มีคอลัมน์ครบแล้ว ไม่ต้อง migrate'); return; }

  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, newCols.length).setValues([newCols]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1 && newCols.indexOf('status') !== -1){
    const statusColIndex = startCol + newCols.indexOf('status');
    const defaults = Array(lastRow - 1).fill(['active']);
    sheet.getRange(2, statusColIndex, lastRow - 1, 1).setValues(defaults);
  }
  Logger.log('เพิ่มคอลัมน์: ' + newCols.join(', '));
}
















function findOrCreateStore(name) {
  const existing = findExactMatch('Stores', name);
  if (existing) return existing;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  const id = generateId('Stores', 'ST');
  sheet.appendRow([id, name.trim(), true]);
  return id;
}

function findOrCreateUnit(name, type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Units');
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const normTarget = normalizeName(name);
    const found = data.find(row => normalizeName(row[1]) === normTarget && row[2] === type);
    if (found) return found[0];
  }

  const id = generateId('Units', 'UNIT');
  sheet.appendRow([id, name.trim(), type]);
  return id;
}


/* ============================================================
   PHASE 3 — บันทึกการซื้อ + คำนวณต้นทุน
   (getPurchasesByProductGroup / getPurchasesByStore เป็นเวอร์ชันอัปเดตจาก Phase 9
    ที่กรองรายการยกเลิกออกให้อัตโนมัติ)
   ============================================================ */

function calcPurchaseDerived(purchase) {
  const qtySellTotal = purchase.qty_buy * purchase.ratio;

  const netCost = purchase.price_buy
    + (purchase.shipping || 0)
    + (purchase.other_cost || 0)
    - (purchase.discount || 0);

  const costPerUnit = qtySellTotal > 0 ? netCost / qtySellTotal : 0;

  return {
    qty_sell_total: qtySellTotal,
    net_cost: netCost,
    cost_per_unit: Math.round(costPerUnit * 100) / 100
  };
}

function savePurchase(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const categoryId = findOrCreateCategory(input.category_name);
   const productGroupId = findOrCreateProductGroup(categoryId, input.product_group_name, input.image_url, input.product_url);
    const storeId = findOrCreateStore(input.store_name);
    findOrCreateUnit(input.unit_buy, 'buy');
    findOrCreateUnit(input.unit_sell, 'sell');

    if (!input.qty_buy || input.qty_buy <= 0) throw new Error('จำนวนซื้อต้องมากกว่า 0');
    if (!input.ratio || input.ratio <= 0) throw new Error('เรโชต้องมากกว่า 0');
    if (input.price_buy == null || input.price_buy < 0) throw new Error('ราคาซื้อไม่ถูกต้อง');

    const discounts = Array.isArray(input.discounts) ? input.discounts.filter(d => d && Number(d.amount) > 0) : [];
    const discountTotal = discounts.reduce((sum, d) => sum + Number(d.amount), 0);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const id = generateId('Purchases', 'PU');

    // เติมคอลัมน์หลัก 18 คอลัมน์แรกตามโครงสร้างเดิมเสมอ (ตำแหน่งคงที่)
    sheet.appendRow([
      id, input.date, productGroupId, storeId, input.order_no || '', input.recorder || '',
      input.note || '', input.qty_buy, input.unit_buy, input.ratio, input.unit_sell,
      input.price_buy, input.shipping || 0, discountTotal, input.other_cost || 0,
      input.receipt_img || '', input.bill_img || '', new Date()
    ]);

    const rowNum = sheet.getLastRow();
    // เติมคอลัมน์เสริม (status / discounts_json) โดยอ้างอิงชื่อ header กันพังไม่ว่าจะอยู่ตำแหน่งไหน
    const statusIdx = headers.indexOf('status');
    if (statusIdx !== -1) sheet.getRange(rowNum, statusIdx + 1).setValue('active');
    const discJsonIdx = headers.indexOf('discounts_json');
    if (discJsonIdx !== -1) sheet.getRange(rowNum, discJsonIdx + 1).setValue(JSON.stringify(discounts));
    const urlIdx = headers.indexOf('product_url');
    if (urlIdx !== -1) sheet.getRange(rowNum, urlIdx + 1).setValue(input.product_url || '');


    return { success: true, id: id, product_group_id: productGroupId };
  } finally {
    lock.releaseLock();
  }
}

function updatePurchase(purchaseId, updates){
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if (!updates.qty_buy || updates.qty_buy <= 0) throw new Error('จำนวนซื้อต้องมากกว่า 0');
    if (!updates.ratio || updates.ratio <= 0) throw new Error('เรโชต้องมากกว่า 0');
    if (updates.price_buy == null || updates.price_buy < 0) throw new Error('ราคาซื้อไม่ถูกต้อง');

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === purchaseId);
    if (rowIndex === -1) throw new Error('ไม่พบรายการซื้อนี้');

    let storeId = data[rowIndex][headers.indexOf('store_id')];
    if (updates.store_name){
      storeId = findOrCreateStore(updates.store_name);
    }
    if (updates.unit_buy) findOrCreateUnit(updates.unit_buy, 'buy');
    if (updates.unit_sell) findOrCreateUnit(updates.unit_sell, 'sell');

    const discounts = Array.isArray(updates.discounts) ? updates.discounts.filter(d => d && Number(d.amount) > 0) : [];
    const discountTotal = discounts.reduce((sum, d) => sum + Number(d.amount), 0);

    const fieldMap = {
  date: updates.date, store_id: storeId, order_no: updates.order_no || '',
  note: updates.note || '', qty_buy: updates.qty_buy, unit_buy: updates.unit_buy,
  ratio: updates.ratio, unit_sell: updates.unit_sell, price_buy: updates.price_buy,
  shipping: updates.shipping || 0, discount: discountTotal, other_cost: updates.other_cost || 0,
  product_url: updates.product_url || ''
};


    const rowNum = rowIndex + 1;
    Object.keys(fieldMap).forEach(field => {
      const colIdx = headers.indexOf(field);
      if (colIdx !== -1 && fieldMap[field] !== undefined){
        sheet.getRange(rowNum, colIdx + 1).setValue(fieldMap[field]);
      }
    });

    const discJsonIdx = headers.indexOf('discounts_json');
    if (discJsonIdx !== -1) sheet.getRange(rowNum, discJsonIdx + 1).setValue(JSON.stringify(discounts));

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}





function getPurchasesByProductGroup(productGroupId, includeCancelled) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const statusIdx = headers.indexOf('status');

  return data
    .filter(row => row[headers.indexOf('product_group_id')] === productGroupId)
    .filter(row => includeCancelled || statusIdx === -1 || row[statusIdx] !== 'cancelled')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      const derived = calcPurchaseDerived(obj);
      return { ...obj, ...derived };
    })
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}

function getPurchasesByStore(productGroupId, storeId, includeCancelled) {
  return getPurchasesByProductGroup(productGroupId, includeCancelled)
    .filter(p => p.store_id === storeId);
}

function getStoreSummary(productGroupId, storeId) {
  const purchases = getPurchasesByStore(productGroupId, storeId);
  if (purchases.length === 0) return null;

  const prices = purchases.map(p => p.cost_per_unit);
  const latest = purchases[0];

  return {
    latest_price: latest.cost_per_unit,
    latest_date: latest.date,
    purchase_count: purchases.length,
    avg_price: Math.round((prices.reduce((a, b) => a + b, 0) / prices.length) * 100) / 100,
    min_price: Math.min(...prices),
    max_price: Math.max(...prices)
  };
}

function getStoreListByProductGroup(productGroupId) {
  const purchases = getPurchasesByProductGroup(productGroupId); // ใหม่ → เก่าอยู่แล้ว

  const storeMap = {};
  purchases.forEach(p => {
    if (!storeMap[p.store_id]) {
      storeMap[p.store_id] = { latest: p, count: 1 }; // รายการแรกที่เจอ (เพราะเรียงใหม่→เก่า) คือรายการล่าสุดของร้านนี้
    } else {
      storeMap[p.store_id].count++;
    }
  });

  const storesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  const storesData = storesSheet.getRange(2, 1, Math.max(storesSheet.getLastRow() - 1, 0), 2).getValues();
  const storeNames = Object.fromEntries(storesData.map(r => [r[0], r[1]]));

  return Object.keys(storeMap)
    .map(storeId => {
      const latest = storeMap[storeId].latest;
      return {
        store_id: storeId,
        store_name: storeNames[storeId] || storeId,
        last_purchase_date: latest.date,
        cost_per_buy_unit: Math.round((latest.net_cost / latest.qty_buy) * 100) / 100,
        cost_per_sell_unit: latest.cost_per_unit,
        unit_buy: latest.unit_buy,
        unit_sell: latest.unit_sell,
        purchase_count: storeMap[storeId].count
      };
    })
    .sort((a, b) => new Date(b.last_purchase_date) - new Date(a.last_purchase_date));
}







/**
 * ดึงการ์ดสินค้า รองรับหลายหมวดพร้อมกันในคำขอเดียว (categoryIds = null คือเอาทุกหมวด)
 * อ่านชีต Purchases แค่ครั้งเดียว แล้ว index ตาม product_group_id ก่อนคำนวณ
 * (ของเดิมอ่านทั้งชีต Purchases ซ้ำทุกสินค้า -> ช้ามากเมื่อประวัติซื้อเยอะ)
 */
function getProductGroupCardsBatch(categoryIds) {
  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgLastRow = pgSheet.getLastRow();
  if (pgLastRow <= 1) return [];

  const pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
  const pgData = pgSheet.getRange(2, 1, pgLastRow - 1, pgSheet.getLastColumn()).getValues();
  const catIdIdx = pgHeaders.indexOf('category_id');
  const activeIdx = pgHeaders.indexOf('active');
  const wantAll = !categoryIds || categoryIds.length === 0;
  const catSet = wantAll ? null : new Set(categoryIds);

  const productGroups = pgData
    .filter(row => row[activeIdx] !== false && (wantAll || catSet.has(row[catIdIdx])))
    .map(row => {
      const pg = {};
      pgHeaders.forEach((h, i) => pg[h] = row[i]);
      return pg;
    });
  if (productGroups.length === 0) return [];

  // อ่านชีต Purchases แค่ครั้งเดียว แล้ว index ตาม product_group_id
  const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const pLastRow = pSheet.getLastRow();
  const purchasesByGroup = {};
  if (pLastRow > 1) {
    const pHeaders = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];
    const pData = pSheet.getRange(2, 1, pLastRow - 1, pSheet.getLastColumn()).getValues();
    const groupIdIdx = pHeaders.indexOf('product_group_id');
    const statusIdx = pHeaders.indexOf('status');

    pData.forEach(row => {
      if (statusIdx !== -1 && row[statusIdx] === 'cancelled') return;
      const gid = row[groupIdIdx];
      const obj = {};
      pHeaders.forEach((h, i) => obj[h] = row[i]);
      const full = { ...obj, ...calcPurchaseDerived(obj) };
      if (!purchasesByGroup[gid]) purchasesByGroup[gid] = [];
      purchasesByGroup[gid].push(full);
    });
    Object.keys(purchasesByGroup).forEach(gid => {
      purchasesByGroup[gid].sort((a, b) => new Date(b.date) - new Date(a.date));
    });
  }

  return productGroups.map(pg => {
    const purchases = purchasesByGroup[pg.id] || [];
    const storeIds = new Set(purchases.map(p => p.store_id));
    const latest = purchases[0];

    const costPerSellUnit = latest ? latest.cost_per_unit : 0;
    const costPerBuyUnit = latest ? Math.round((latest.net_cost / latest.qty_buy) * 100) / 100 : 0;
    const unitBuy = latest ? latest.unit_buy : '';
    const unitSell = latest ? latest.unit_sell : '';
    const ratio = latest ? latest.ratio : 0;

    const sellPricePerSellUnit = pg.current_sell_price || 0;
    const sellPricePerBuyUnit = ratio ? Math.round(sellPricePerSellUnit * ratio * 100) / 100 : 0;

      let purchaseFrequencyDays = null;
    if (purchases.length >= 2) {
      const datesAsc = purchases.map(p => new Date(p.date)).sort((a, b) => a - b);
      let totalDays = 0;
      for (let i = 1; i < datesAsc.length; i++) totalDays += (datesAsc[i] - datesAsc[i - 1]) / 86400000;
      purchaseFrequencyDays = Math.round(totalDays / (datesAsc.length - 1));
    }
    const daysSinceLastPurchase = latest ? Math.floor((new Date() - new Date(latest.date)) / 86400000) : null;
    const shouldReorder = purchaseFrequencyDays !== null && daysSinceLastPurchase !== null && daysSinceLastPurchase >= purchaseFrequencyDays;

    return {
      id: pg.id, name: pg.name, image_url: pg.image_url,
      last_purchase_date: latest ? latest.date : null,
      cost_per_buy_unit: costPerBuyUnit, unit_buy: unitBuy,
      cost_per_sell_unit: costPerSellUnit, unit_sell: unitSell,
      sell_price_per_buy_unit: sellPricePerBuyUnit,
      sell_price_per_sell_unit: sellPricePerSellUnit,
      store_count: storeIds.size,
      purchase_count: purchases.length,
      purchase_frequency_days: purchaseFrequencyDays,
      days_since_last_purchase: daysSinceLastPurchase,
      should_reorder: shouldReorder
    };

  });
}



/* ============================================================
   PHASE 4 — ตั้งราคาขาย + ประวัติราคา
   ============================================================ */

function getPricingInfo(productGroupId) {
  const purchases = getPurchasesByProductGroup(productGroupId);
  if (purchases.length === 0) throw new Error('ยังไม่มีประวัติการซื้อของกลุ่มสินค้านี้');

  const costPerUnit = purchases[0].cost_per_unit;

  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const lastRow = pgSheet.getLastRow();
  const headers = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
  const data = pgSheet.getRange(2, 1, lastRow - 1, pgSheet.getLastColumn()).getValues();

  const idIdx = headers.indexOf('id');
  const priceIdx = headers.indexOf('current_sell_price');
  const row = data.find(r => r[idIdx] === productGroupId);
  if (!row) throw new Error('ไม่พบกลุ่มสินค้านี้');

  const currentSellPrice = row[priceIdx] || 0;
  const profit = currentSellPrice - costPerUnit;
  const margin = currentSellPrice > 0 ? (profit / currentSellPrice) * 100 : 0;

  return {
    product_group_id: productGroupId, cost_per_unit: costPerUnit, current_sell_price: currentSellPrice,
    profit_per_unit: Math.round(profit * 100) / 100, margin_percent: Math.round(margin * 10) / 10,
    status: getProfitStatus(margin)
  };
}

function getProfitStatus(marginPercent) {
  if (marginPercent < 0) return { emoji: '🔴', label: 'ขาดทุน' };
  if (marginPercent < 15) return { emoji: '🟡', label: 'ใกล้ทุน' };
  return { emoji: '🟢', label: 'กำไร' };
}

function calcFromSellPrice(costPerUnit, sellPrice) {
  const profit = sellPrice - costPerUnit;
  const margin = sellPrice > 0 ? (profit / sellPrice) * 100 : 0;
  return { sell_price: sellPrice, profit_per_unit: Math.round(profit * 100) / 100, margin_percent: Math.round(margin * 10) / 10 };
}

function calcFromMarginPercent(costPerUnit, marginPercent, roundToInteger) {
  let sellPrice = costPerUnit / (1 - marginPercent / 100);
  sellPrice = roundToInteger ? Math.ceil(sellPrice) : Math.round(sellPrice * 100) / 100;
  const profit = sellPrice - costPerUnit;
  return { sell_price: sellPrice, profit_per_unit: Math.round(profit * 100) / 100, margin_percent: marginPercent };
}


function saveSellPrice(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const info = getPricingInfo(input.product_group_id);
    const oldPrice = info.current_sell_price;
    const costPerUnit = info.cost_per_unit;

    const profit = input.new_price - costPerUnit;
    const margin = input.new_price > 0 ? (profit / input.new_price) * 100 : 0;

    const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
    const lastRow = pgSheet.getLastRow();
    const headers = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
    const idIdx = headers.indexOf('id');
    const priceIdx = headers.indexOf('current_sell_price');

    const data = pgSheet.getRange(2, 1, lastRow - 1, pgSheet.getLastColumn()).getValues();
    const rowIndex = data.findIndex(r => r[idIdx] === input.product_group_id);
    if (rowIndex === -1) throw new Error('ไม่พบกลุ่มสินค้านี้');

    pgSheet.getRange(rowIndex + 2, priceIdx + 1).setValue(input.new_price);

    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PriceHistory');
    const id = generateId('PriceHistory', 'PH');
    historySheet.appendRow([
      id, input.product_group_id, new Date(), costPerUnit, oldPrice, input.new_price,
      Math.round(profit * 100) / 100, Math.round(margin * 10) / 10, input.edited_by || ''
    ]);

    return {
      success: true, old_price: oldPrice, new_price: input.new_price,
      profit_per_unit: Math.round(profit * 100) / 100, margin_percent: Math.round(margin * 10) / 10
    };
  } finally {
    lock.releaseLock();
  }
}

function getPriceHistory(productGroupId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PriceHistory');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const pgIdx = headers.indexOf('product_group_id');

  return data
    .filter(row => row[pgIdx] === productGroupId)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    })
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}


/* ============================================================
   PHASE 5 — ฟังก์ชันรวมข้อมูลสำหรับแต่ละหน้า
   (getStoreDetail เป็นเวอร์ชันอัปเดตจาก Phase 9 ที่รองรับ includeCancelled)
   ============================================================ */

function getCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data.filter(r => r[2] !== false).map(r => ({ id: r[0], name: r[1] }));
}

function getAllStores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data.filter(r => r[2] !== false).map(r => ({ id: r[0], name: r[1] }));
}


/**5) เพิ่มฟังก์ชันใหม่ (สำหรับปุ่ม ✏️ แก้ไขลิงก์ทีหลัง)**
 * 
 * 
 */
function updateProductGroupLink(productGroupId, url){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const urlIdx = headers.indexOf('product_url');
  if (urlIdx === -1) throw new Error('กรุณารัน migrateProductUrlColumn() ก่อน');
  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === productGroupId);
  if (rowIndex === -1) throw new Error('ไม่พบกลุ่มสินค้านี้');
  sheet.getRange(rowIndex + 1, urlIdx + 1).setValue(url || '');
  return { success: true };
}







function getStoreDetail(productGroupId, storeId, includeCancelled) {
  const summary = getStoreSummary(productGroupId, storeId);
  const purchases = getPurchasesByStore(productGroupId, storeId, includeCancelled);

  const storesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  const storeRow = storesSheet.getDataRange().getValues().find(r => r[0] === storeId);

  return {
    store_id: storeId,
    store_name: storeRow ? storeRow[1] : storeId,
    summary: summary,
    purchases: purchases.map(p => ({
      id: p.id, date: p.date, price_buy: p.price_buy, cost_per_unit: p.cost_per_unit,
      qty_buy: p.qty_buy, unit_buy: p.unit_buy, status: p.status || 'active'
    }))
  };
}

function getPurchaseDetail(purchaseId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const row = data.find(r => r[0] === purchaseId);
  if (!row) throw new Error('ไม่พบรายการซื้อนี้');

  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  const derived = calcPurchaseDerived(obj);

  let discounts = [];
  try { discounts = obj.discounts_json ? JSON.parse(obj.discounts_json) : []; } catch (err) { discounts = []; }
  if (discounts.length === 0 && obj.discount > 0) {
    discounts = [{ name: 'ส่วนลด', amount: obj.discount }]; // รองรับรายการเก่าก่อน migrate
  }

  const storesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  const storeRow = storesSheet.getDataRange().getValues().find(r => r[0] === obj.store_id);
  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgRow = pgSheet.getDataRange().getValues().find(r => r[0] === obj.product_group_id);

  return {
    ...obj, ...derived, discounts,
    store_name: storeRow ? storeRow[1] : obj.store_id,
    product_group_name: pgRow ? pgRow[2] : obj.product_group_id,
    image_url: pgRow ? pgRow[3] : ''
  };
}





function searchMasterFor(type, keyword, categoryId) {
  const sheetMap = { category: 'Categories', productGroup: 'ProductGroups', store: 'Stores' };
  const sheetName = sheetMap[type];
  if (!sheetName) return [];
  let results = searchMaster(sheetName, keyword || '');
  if (type === 'productGroup' && categoryId) {
    const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
    const data = pgSheet.getDataRange().getValues();
    const validIds = new Set(data.filter(r => r[1] === categoryId).map(r => r[0]));
    results = results.filter(r => validIds.has(r.id));
  }
  return results;
}


/* ============================================================
   PHASE 6 — ⚙️ จัดการข้อมูล (แก้ไข / ปิดใช้งาน / ลบ / รวม)
   ============================================================ */

const MASTER_CONFIG = {
  category:     { sheet: 'Categories',    refs: [ { sheet: 'ProductGroups', field: 'category_id' } ] },
  productGroup: { sheet: 'ProductGroups', refs: [ { sheet: 'Purchases', field: 'product_group_id' },
                                                   { sheet: 'PriceHistory', field: 'product_group_id' } ] },
  store:        { sheet: 'Stores',        refs: [ { sheet: 'Purchases', field: 'store_id' } ] }
};

function _sheetData(sheetName){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { sheet, headers, data };
}
function _col(headers, name){ return headers.indexOf(name); }

function getMasterList(type, categoryId){
  if (type === 'unit_buy' || type === 'unit_sell'){
    const unitType = type === 'unit_buy' ? 'buy' : 'sell';
    const { headers, data } = _sheetData('Units');
    const idIdx = _col(headers, 'id'), nameIdx = _col(headers, 'name'), typeIdx = _col(headers, 'type');
    const purchases = _sheetData('Purchases');
    const field = unitType === 'buy' ? 'unit_buy' : 'unit_sell';
    const pFieldIdx = _col(purchases.headers, field);

    // นับการใช้งานล่วงหน้าครั้งเดียว แทนการ filter ทั้งชีตซ้ำทุกแถว (ของเดิมทำให้ช้ามากเมื่อข้อมูลเยอะ)
    const usageCount = {};
    purchases.data.forEach(p => {
      const val = p[pFieldIdx];
      usageCount[val] = (usageCount[val] || 0) + 1;
    });

    return data
      .filter(r => r[typeIdx] === unitType)
      .map(r => ({ id: r[idIdx], name: r[nameIdx], active: true, usage: usageCount[r[nameIdx]] || 0 }))
      .sort((a, b) => a.name.localeCompare(b.name, 'th'));
  }

  const cfg = MASTER_CONFIG[type];
  if (!cfg) throw new Error('ประเภทข้อมูลไม่ถูกต้อง');
  const { headers, data } = _sheetData(cfg.sheet);
  const idIdx = _col(headers, 'id'), nameIdx = _col(headers, 'name'), activeIdx = _col(headers, 'active');
  const catIdx = type === 'productGroup' ? _col(headers, 'category_id') : -1;

  // ⚠️ จุดที่แก้: อ่านชีตอ้างอิงแต่ละอันแค่ "ครั้งเดียว" แล้วนับรวมไว้ล่วงหน้า
  // ของเดิมอ่านทั้งชีต Purchases/PriceHistory ซ้ำใหม่ทุกแถวของ ProductGroups -> ช้ามาก (O(N×M))
  const usageCount = {};
  cfg.refs.forEach(ref => {
    const refData = _sheetData(ref.sheet);
    const fIdx = _col(refData.headers, ref.field);
    refData.data.forEach(rr => {
      const val = rr[fIdx];
      usageCount[val] = (usageCount[val] || 0) + 1;
    });
  });

  let rows = data;
  if (type === 'productGroup' && categoryId){
    rows = rows.filter(r => r[catIdx] === categoryId); // กรองตามหมวดที่เลือก (ใช้สำหรับข้อ 1)
  }

  return rows
    .map(r => ({
      id: r[idIdx], name: r[nameIdx],
      active: activeIdx !== -1 ? r[activeIdx] !== false : true,
      usage: usageCount[r[idIdx]] || 0,
      category_id: catIdx !== -1 ? r[catIdx] : undefined
    }))
    .sort((a, b) => a.name.localeCompare(b.name, 'th'));
}


function renameMaster(type, id, newName){
  newName = String(newName).trim();
  if (!newName) throw new Error('กรุณากรอกชื่อ');

  if (type === 'unit_buy' || type === 'unit_sell'){
    const unitType = type === 'unit_buy' ? 'buy' : 'sell';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Units');
    const { headers, data } = _sheetData('Units');
    const idIdx = _col(headers, 'id'), nameIdx = _col(headers, 'name'), typeIdx = _col(headers, 'type');

    const dup = data.find(r => r[typeIdx] === unitType && normalizeName(r[nameIdx]) === normalizeName(newName) && r[idIdx] !== id);
    if (dup) throw new Error('มีชื่อนี้อยู่แล้ว — กรุณาใช้ปุ่ม "รวม" แทน');

    const rowIndex = data.findIndex(r => r[idIdx] === id);
    if (rowIndex === -1) throw new Error('ไม่พบข้อมูล');
    const oldName = data[rowIndex][nameIdx];
    sheet.getRange(rowIndex + 2, nameIdx + 1).setValue(newName);

    const field = unitType === 'buy' ? 'unit_buy' : 'unit_sell';
    const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
    const pData = _sheetData('Purchases');
    const pFieldIdx = _col(pData.headers, field);
    pData.data.forEach((r, i) => {
      if (r[pFieldIdx] === oldName) pSheet.getRange(i + 2, pFieldIdx + 1).setValue(newName);
    });
    return { success: true };
  }

  const cfg = MASTER_CONFIG[type];
  if (!cfg) throw new Error('ประเภทข้อมูลไม่ถูกต้อง');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
  const { headers, data } = _sheetData(cfg.sheet);
  const idIdx = _col(headers, 'id'), nameIdx = _col(headers, 'name');
  const rowIndex = data.findIndex(r => r[idIdx] === id);
  if (rowIndex === -1) throw new Error('ไม่พบข้อมูล');

  let dup;
  if (type === 'productGroup'){
    const catIdx = _col(headers, 'category_id');
    const catId = data[rowIndex][catIdx];
    dup = data.find(r => r[catIdx] === catId && normalizeName(r[nameIdx]) === normalizeName(newName) && r[idIdx] !== id);
  } else {
    dup = data.find(r => normalizeName(r[nameIdx]) === normalizeName(newName) && r[idIdx] !== id);
  }
  if (dup) throw new Error('มีชื่อนี้อยู่แล้ว — กรุณาใช้ปุ่ม "รวม" แทน');

  sheet.getRange(rowIndex + 2, nameIdx + 1).setValue(newName);
  return { success: true };
}

function setMasterActive(type, id, active){
  const cfg = MASTER_CONFIG[type];
  if (!cfg) throw new Error('ประเภทนี้ไม่รองรับการปิดใช้งาน ใช้ลบแทน');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
  const { headers, data } = _sheetData(cfg.sheet);
  const idIdx = _col(headers, 'id'), activeIdx = _col(headers, 'active');
  const rowIndex = data.findIndex(r => r[idIdx] === id);
  if (rowIndex === -1) throw new Error('ไม่พบข้อมูล');
  sheet.getRange(rowIndex + 2, activeIdx + 1).setValue(active);
  return { success: true };
}

function deleteMasterPermanent(type, id){
  const item = getMasterList(type).find(x => x.id === id);
  if (!item) throw new Error('ไม่พบข้อมูล');
  if (item.usage > 0) throw new Error('ยังมีข้อมูลอ้างอิงอยู่ ลบไม่ได้ — ให้ปิดใช้งาน หรือรวมกับรายการอื่นแทน');

  const sheetName = (type === 'unit_buy' || type === 'unit_sell') ? 'Units' : MASTER_CONFIG[type].sheet;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const { headers, data } = _sheetData(sheetName);
  const idIdx = _col(headers, 'id');
  const rowIndex = data.findIndex(r => r[idIdx] === id);
  if (rowIndex === -1) throw new Error('ไม่พบข้อมูล');
  sheet.deleteRow(rowIndex + 2);
  return { success: true };
}

function mergeMaster(type, sourceIds, targetId){
  sourceIds = sourceIds.filter(id => id !== targetId);
  if (sourceIds.length === 0) throw new Error('ไม่มีรายการให้รวม');

  if (type === 'unit_buy' || type === 'unit_sell'){
    const unitType = type === 'unit_buy' ? 'buy' : 'sell';
    const { headers, data } = _sheetData('Units');
    const idIdx = _col(headers, 'id'), nameIdx = _col(headers, 'name');
    const targetRow = data.find(r => r[idIdx] === targetId);
    if (!targetRow) throw new Error('ไม่พบรายการปลายทาง');
    const targetName = targetRow[nameIdx];

    const field = unitType === 'buy' ? 'unit_buy' : 'unit_sell';
    const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
    const pData = _sheetData('Purchases');
    const pFieldIdx = _col(pData.headers, field);

    sourceIds.forEach(sid => {
      const srcRow = data.find(r => r[idIdx] === sid);
      if (!srcRow) return;
      const srcName = srcRow[nameIdx];
      pData.data.forEach((r, i) => {
        if (r[pFieldIdx] === srcName) pSheet.getRange(i + 2, pFieldIdx + 1).setValue(targetName);
      });
    });

    const unitSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Units');
    const rowsToDelete = [];
    data.forEach((r, i) => { if (sourceIds.includes(r[idIdx])) rowsToDelete.push(i + 2); });
    rowsToDelete.sort((a, b) => b - a).forEach(rowNum => unitSheet.deleteRow(rowNum));
    return { success: true, merged: sourceIds.length };
  }

  const cfg = MASTER_CONFIG[type];
  if (!cfg) throw new Error('ประเภทข้อมูลไม่ถูกต้อง');

  cfg.refs.forEach(ref => {
    const refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ref.sheet);
    const { headers, data } = _sheetData(ref.sheet);
    const fIdx = _col(headers, ref.field);
    data.forEach((r, i) => {
      if (sourceIds.includes(r[fIdx])) refSheet.getRange(i + 2, fIdx + 1).setValue(targetId);
    });
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
  const { headers, data } = _sheetData(cfg.sheet);
  const idIdx = _col(headers, 'id');
  const rowsToDelete = [];
  data.forEach((r, i) => { if (sourceIds.includes(r[idIdx])) rowsToDelete.push(i + 2); });
  rowsToDelete.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));

  return { success: true, merged: sourceIds.length };
}


/* ============================================================
   เพิ่มรายการใหม่ (หมวดสินค้า / ร้านค้า / หน่วย) + ย้ายกลุ่มสินค้าไปหมวดอื่นทีละหลายรายการ
   วางต่อท้าย mergeMaster() ในไฟล์ Code.gs (อยู่ในโซนเดียวกับฟังก์ชัน master data อื่นๆ)
   ============================================================ */

function addMaster(type, name){
  name = String(name).trim();
  if (!name) throw new Error('กรุณากรอกชื่อ');

  if (type === 'category'){
    if (findExactMatch('Categories', name)) throw new Error('มีหมวดสินค้านี้อยู่แล้ว');
    const id = findOrCreateCategory(name);
    return { success: true, id };
  }

  if (type === 'store'){
    if (findExactMatch('Stores', name)) throw new Error('มีร้านค้านี้อยู่แล้ว');
    const id = findOrCreateStore(name);
    return { success: true, id };
  }

  if (type === 'unit_buy' || type === 'unit_sell'){
    const unitType = type === 'unit_buy' ? 'buy' : 'sell';
    const { headers, data } = _sheetData('Units');
    const nameIdx = _col(headers, 'name'), typeIdx = _col(headers, 'type');
    const dup = data.find(r => r[typeIdx] === unitType && normalizeName(r[nameIdx]) === normalizeName(name));
    if (dup) throw new Error('มีหน่วยนี้อยู่แล้ว');
    const id = findOrCreateUnit(name, unitType);
    return { success: true, id };
  }

  throw new Error('ประเภทนี้เพิ่มรายการใหม่ตรงนี้ไม่ได้ (กลุ่มสินค้าจะถูกสร้างอัตโนมัติตอนบันทึกการซื้อ)');
}

/**
 * ย้ายกลุ่มสินค้า (ProductGroups) หลายรายการไปหมวดสินค้าปลายทางในคราวเดียว
 * ข้ามรายการที่ชื่อชนกับกลุ่มสินค้าที่มีอยู่แล้วในหมวดปลายทาง (กันข้อมูลซ้ำ)
 */
function moveProductGroupsToCategory(ids, targetCategoryId){
  if (!ids || ids.length === 0) throw new Error('ไม่มีรายการที่เลือก');
  if (!targetCategoryId) throw new Error('กรุณาเลือกหมวดปลายทาง');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const { headers, data } = _sheetData('ProductGroups');
  const idIdx = _col(headers, 'id'), catIdx = _col(headers, 'category_id'), nameIdx = _col(headers, 'name');

  // กันชื่อชนกับกลุ่มสินค้าที่มีอยู่แล้วในหมวดปลายทาง
  const targetNames = new Set(
    data.filter(r => r[catIdx] === targetCategoryId).map(r => normalizeName(r[nameIdx]))
  );

  let moved = 0, skipped = 0;
  ids.forEach(id => {
    const rowIndex = data.findIndex(r => r[idIdx] === id);
    if (rowIndex === -1) return;
    if (data[rowIndex][catIdx] === targetCategoryId) return; // อยู่หมวดเดียวกันอยู่แล้ว ข้าม

    const name = data[rowIndex][nameIdx];
    if (targetNames.has(normalizeName(name))) { skipped++; return; }

    sheet.getRange(rowIndex + 2, catIdx + 1).setValue(targetCategoryId);
    targetNames.add(normalizeName(name));
    moved++;
  });

  return { success: true, moved, skipped };
}




/* ============================================================
   PHASE 7 — อัปโหลดรูปกลุ่มสินค้า
   ============================================================ */

function _getProductImageFolder(){
  const folderName = 'ProductImages_ระบบจัดการสินค้า';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function uploadProductGroupImage(productGroupId, base64Data, mimeType, fileName){
  if (!base64Data) throw new Error('ไม่พบข้อมูลรูปภาพ');

  const folder = _getProductImageFolder();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', fileName || (productGroupId + '.jpg'));

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const imgIdx = headers.indexOf('image_url');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === productGroupId);
  if (rowIndex === -1) throw new Error('ไม่พบกลุ่มสินค้านี้');

  const oldUrl = data[rowIndex][imgIdx];
  if (oldUrl && oldUrl.indexOf('drive.google.com') !== -1){
    try {
      const match = oldUrl.match(/id=([a-zA-Z0-9_-]+)/);
      if (match) DriveApp.getFileById(match[1]).setTrashed(true);
    } catch (err) { /* ไฟล์เก่าอาจถูกลบไปแล้ว ข้ามได้ */ }
  }

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';

  sheet.getRange(rowIndex + 1, imgIdx + 1).setValue(url);
  return { success: true, url: url };
}


function removeProductGroupImage(productGroupId){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const imgIdx = headers.indexOf('image_url');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === productGroupId);
  if (rowIndex === -1) throw new Error('ไม่พบกลุ่มสินค้านี้');

  const oldUrl = data[rowIndex][imgIdx];
  if (oldUrl && oldUrl.indexOf('drive.google.com') !== -1){
    try {
      const match = oldUrl.match(/id=([a-zA-Z0-9_-]+)/);
      if (match) DriveApp.getFileById(match[1]).setTrashed(true);
    } catch (err) { /* ข้ามได้ */ }
  }

  sheet.getRange(rowIndex + 1, imgIdx + 1).setValue('');
  return { success: true };
}


/* ============================================================
   PHASE 8 — 📊 ประวัติราคา
   ============================================================ */

function getPriceHistoryDetail(productGroupId){
  const history = getPriceHistory(productGroupId);

  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgRow = pgSheet.getDataRange().getValues().find(r => r[0] === productGroupId);
  if (!pgRow) throw new Error('ไม่พบกลุ่มสินค้านี้');

  return {
    product_group_id: productGroupId, name: pgRow[2], image_url: pgRow[3],
    history: history.map(h => ({
      date: h.date, cost_price: h.cost_price, old_price: h.old_price,
      new_price: h.new_price, profit: h.profit, margin: h.margin, edited_by: h.edited_by
    }))
  };
}








function cancelPurchase(purchaseId, reason){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === purchaseId);
  if (rowIndex === -1) throw new Error('ไม่พบรายการซื้อนี้');

  const statusIdx = headers.indexOf('status');
  const reasonIdx = headers.indexOf('cancel_reason');
  const cancelledAtIdx = headers.indexOf('cancelled_at');
  if (statusIdx === -1) throw new Error('กรุณารัน migratePurchaseStatusColumn() ก่อนใช้งานฟีเจอร์นี้');

  const rowNum = rowIndex + 1;
  sheet.getRange(rowNum, statusIdx + 1).setValue('cancelled');
  sheet.getRange(rowNum, reasonIdx + 1).setValue(reason || '');
  sheet.getRange(rowNum, cancelledAtIdx + 1).setValue(new Date());
  return { success: true };
}

function restorePurchase(purchaseId){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === purchaseId);
  if (rowIndex === -1) throw new Error('ไม่พบรายการซื้อนี้');

  const statusIdx = headers.indexOf('status');
  const reasonIdx = headers.indexOf('cancel_reason');
  const cancelledAtIdx = headers.indexOf('cancelled_at');
  const rowNum = rowIndex + 1;
  sheet.getRange(rowNum, statusIdx + 1).setValue('active');
  sheet.getRange(rowNum, reasonIdx + 1).setValue('');
  sheet.getRange(rowNum, cancelledAtIdx + 1).setValue('');
  return { success: true };
}


/* ============================================================
   PHASE 10 — จำกัดสิทธิ์เข้าถึง + doGet/include (จุดเริ่มต้นเว็บแอป)
   ============================================================ */

function _getAllowedSheet(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AllowedUsers');
  if (!sheet){
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('AllowedUsers');
    sheet.getRange(1, 1, 1, 3).setValues([['email', 'name', 'active']]);
    let ownerEmail = '';
    try { ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(); } catch (err) {}
    if (!ownerEmail) ownerEmail = Session.getEffectiveUser().getEmail();
    if (ownerEmail) sheet.appendRow([ownerEmail, 'เจ้าของระบบ', true]);
  }
  return sheet;
}

function isEmailAllowed(email){
  if (!email) return false;
  const sheet = _getAllowedSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const target = email.toString().toLowerCase().trim();
  return data.some(r => r[0] && r[0].toString().toLowerCase().trim() === target && r[2] !== false);
}

function getAllowedUsers(){
  const sheet = _getAllowedSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .map((r, i) => ({ row: i + 2, email: r[0], name: r[1], active: r[2] !== false }));
}

function addAllowedUser(email, name){
  email = String(email).trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) throw new Error('กรุณากรอกอีเมลให้ถูกต้อง');
  const existing = getAllowedUsers().find(u => u.email.toLowerCase() === email);
  if (existing) throw new Error('มีอีเมลนี้ในระบบอยู่แล้ว');
  _getAllowedSheet().appendRow([email, name || '', true]);
  return { success: true };
}

function setAllowedUserActive(rowNum, active){
  _getAllowedSheet().getRange(rowNum, 3).setValue(active);
  return { success: true };
}

function removeAllowedUser(rowNum){
  const users = getAllowedUsers();
  const activeCount = users.filter(u => u.active).length;
  if (activeCount <= 1) throw new Error('ต้องมีผู้ใช้ที่เปิดใช้งานอยู่อย่างน้อย 1 คน กันล็อกตัวเองออกจากระบบ');
  _getAllowedSheet().deleteRow(rowNum);
  return { success: true };
}

function getCurrentUserEmail(){
  return Session.getActiveUser().getEmail() || '';
}

/** จุดเริ่มต้นเว็บแอป — ตรวจสิทธิ์ก่อนเสิร์ฟหน้า Index */
function doGet(e){
  const email = Session.getActiveUser().getEmail();

  if (!isEmailAllowed(email)){
    return HtmlService.createTemplateFromFile('AccessDenied')
      .evaluate()
      .setTitle('ไม่มีสิทธิ์เข้าถึง')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  // ส่งพารามิเตอร์จาก query string เข้าไปในหน้า เผื่อ #hash หลุดหายระหว่าง redirect ของ Apps Script
  template.initialPage = e.parameter.page || '';
  template.initialId = e.parameter.pgId || '';

  return template
    .evaluate()
    .setTitle('ระบบจัดการสินค้า')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ============================================================
 * JSON BRIDGE — แก้บั๊ก google.script.run คืนค่า null แบบเงียบๆ
 * เพิ่มต่อท้ายไฟล์ Code.gs
 *
 * สาเหตุ: google.script.run มีบั๊กที่รู้จักกันดี — เวลา server คืนค่า
 * เป็น object ซับซ้อน (มี array ซ้อน, Date object) การส่งผ่าน bridge
 * ภายในของมันเองบางทีพังเงียบๆ กลายเป็น null ที่ฝั่ง client
 * โดยไม่มี error ให้เห็นเลย (ตรงกับที่เจอ)
 *
 * ทางแก้: บีบทุกอย่างให้เป็น JSON string ก่อนส่ง แล้วแปลงกลับฝั่ง client
 * ============================================================
 */
function callServer(fnName, argsJsonStr){
  const args = argsJsonStr ? JSON.parse(argsJsonStr) : [];
  const fn = this[fnName];
  if (typeof fn !== 'function') throw new Error('ไม่พบฟังก์ชัน: ' + fnName);
  const result = fn.apply(null, args);
  return JSON.stringify(result === undefined ? null : result);
}


/**
 * ============================================================
 * ส่วนลดแบบยืดหยุ่น — รองรับหลายรายการ ชื่อไม่ตายตัว
 * เพิ่มต่อท้ายไฟล์ Code.gs
 *
 * แนวคิด: เพิ่มคอลัมน์ discounts_json เก็บเป็น JSON array
 * เช่น [{"name":"ส่วนลดสมาชิก","amount":20},{"name":"โปรโมชั่น","amount":15}]
 * ส่วนคอลัมน์ discount เดิม (ตัวเลขรวม) ยังเก็บผลรวมไว้เหมือนเดิม
 * เพื่อให้สูตรคำนวณต้นทุนสุทธิเดิม (calcPurchaseDerived) ไม่ต้องแก้อะไรเลย
 *
 * ขั้นตอนติดตั้ง:
 *   1. วางไฟล์นี้ต่อท้าย Code.gs
 *   2. รัน migrateDiscountsColumn() ครั้งเดียว
 *   3. แทนที่ 3 ฟังก์ชันเดิม (savePurchase, updatePurchase, getPurchaseDetail)
 *      ด้วยเวอร์ชันในคอมเมนต์ REPLACE ท้ายไฟล์นี้
 * ============================================================
 */

function migrateDiscountsColumn(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (headers.indexOf('discounts_json') !== -1){
    Logger.log('มีคอลัมน์ discounts_json อยู่แล้ว ไม่ต้อง migrate');
    return;
  }

  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, 1).setValues([['discounts_json']]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1){
    const discountIdx = headers.indexOf('discount');
    const oldValues = sheet.getRange(2, discountIdx + 1, lastRow - 1, 1).getValues();
    const jsonValues = oldValues.map(r => {
      const amt = Number(r[0]) || 0;
      return [ amt > 0 ? JSON.stringify([{ name: 'ส่วนลด', amount: amt }]) : '[]' ];
    });
    sheet.getRange(2, startCol, lastRow - 1, 1).setValues(jsonValues);
  }
  Logger.log('เพิ่มคอลัมน์ discounts_json แล้ว (แปลงส่วนลดเก่าเป็นรายการเดียวให้อัตโนมัติ)');
}




/**
 * ============================================================
 * รูปใบเสร็จ/บิล — เพิ่มต่อท้ายไฟล์ Code.gs
 * ใช้แนวทางเดียวกับการอัปโหลดรูปกลุ่มสินค้า (Phase 7) แต่แยกโฟลเดอร์
 * ============================================================
 */

function _getPurchaseImageFolder(){
  const folderName = 'PurchaseReceipts_ระบบจัดการสินค้า';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}


/**
 * purchaseId: string
 * imageType: 'receipt_img' หรือ 'bill_img'
 * base64Data: base64 ล้วนๆ (ไม่มี prefix data:image/...)
 */
function uploadPurchaseImage(purchaseId, imageType, base64Data, mimeType, fileName){
  if (imageType !== 'receipt_img' && imageType !== 'bill_img') throw new Error('ประเภทรูปไม่ถูกต้อง');
  if (!base64Data) throw new Error('ไม่พบข้อมูลรูปภาพ');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const imgIdx = headers.indexOf(imageType);
  if (imgIdx === -1) throw new Error('ไม่พบคอลัมน์ ' + imageType);

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === purchaseId);
  if (rowIndex === -1) throw new Error('ไม่พบรายการซื้อนี้');

  const oldUrl = data[rowIndex][imgIdx];
  if (oldUrl && oldUrl.indexOf('drive.google.com') !== -1){
    try {
      const match = oldUrl.match(/id=([a-zA-Z0-9_-]+)/);
      if (match) DriveApp.getFileById(match[1]).setTrashed(true);
    } catch (err) { /* ข้ามได้ */ }
  }

  const folder = _getPurchaseImageFolder();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', fileName || (purchaseId + '_' + imageType + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';

  sheet.getRange(rowIndex + 1, imgIdx + 1).setValue(url);
  return { success: true, url: url };
}

/**
 * ใช้ URL รูปที่อัปโหลดไปแล้ว (จากรายการอื่นในออเดอร์เดียวกัน) มาผูกกับ purchase นี้
 * ไม่สร้างไฟล์ใหม่บน Drive — กันไฟล์ซ้ำซ้อนเวลาออเดอร์เดียวมีหลายสินค้าแต่ใช้รูปใบเดียวกัน
 */
function setPurchaseImageUrl(purchaseId, imageType, url){
  if (imageType !== 'receipt_img' && imageType !== 'bill_img') throw new Error('ประเภทรูปไม่ถูกต้อง');
  if (!url) throw new Error('ไม่พบ URL รูปภาพ');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const imgIdx = headers.indexOf(imageType);
  if (imgIdx === -1) throw new Error('ไม่พบคอลัมน์ ' + imageType);

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === purchaseId);
  if (rowIndex === -1) throw new Error('ไม่พบรายการซื้อนี้');

  sheet.getRange(rowIndex + 1, imgIdx + 1).setValue(url);
  return { success: true, url: url };
}



function removePurchaseImage(purchaseId, imageType){
  if (imageType !== 'receipt_img' && imageType !== 'bill_img') throw new Error('ประเภทรูปไม่ถูกต้อง');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const imgIdx = headers.indexOf(imageType);

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === purchaseId);
  if (rowIndex === -1) throw new Error('ไม่พบรายการซื้อนี้');

  const oldUrl = data[rowIndex][imgIdx];
  if (oldUrl && oldUrl.indexOf('drive.google.com') !== -1){
    try {
      const match = oldUrl.match(/id=([a-zA-Z0-9_-]+)/);
      if (match) DriveApp.getFileById(match[1]).setTrashed(true);
    } catch (err) { /* ข้ามได้ */ }
  }

  sheet.getRange(rowIndex + 1, imgIdx + 1).setValue('');
  return { success: true };
}

/**
 * ============================================================
 * แก้ไข: รูปที่อัปโหลดไม่แสดงผล + ปัดราคาขึ้นเป็นจำนวนเต็ม
 * เพิ่มต่อท้ายไฟล์ Code.gs
 *
 * สาเหตุที่รูปไม่ขึ้น: URL รูปแบบเดิม
 *   https://drive.google.com/uc?export=view&id=...
 * ใช้แสดงในแท็บ <img> ของหน้าเว็บไม่ได้เสถียร (Google บล็อก/รีไดเรกต์บ่อย)
 * เปลี่ยนเป็นรูปแบบ thumbnail ที่เสถียรกว่า:
 *   https://drive.google.com/thumbnail?id=...&sz=w1000
 *
 * ขั้นตอนติดตั้ง:
 *   1. วางไฟล์นี้ต่อท้าย Code.gs
 *   2. รัน fixExistingImageUrls() ครั้งเดียว (แก้ URL รูปเก่าที่อัปโหลดไปแล้วให้ใช้งานได้)
 *   3. แทนที่ uploadProductGroupImage(), uploadPurchaseImage(), calcFromMarginPercent()
 *      ด้วยเวอร์ชันในคอมเมนต์ REPLACE ท้ายไฟล์นี้
 * ============================================================
 */

/** รันครั้งเดียว: แก้ URL รูปเก่าที่เป็นรูปแบบพังให้เป็นรูปแบบใหม่ */
function fixExistingImageUrls(){
  let fixedCount = 0;

  function fixUrl(url){
    if (!url || url.indexOf('drive.google.com') === -1) return url;
    const match = url.match(/id=([a-zA-Z0-9_-]+)/);
    if (!match) return url;
    return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w1000';
  }

  // ProductGroups.image_url
  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgData = pgSheet.getDataRange().getValues();
  const pgHeaders = pgData[0];
  const imgIdx = pgHeaders.indexOf('image_url');
  for (let i = 1; i < pgData.length; i++){
    const oldUrl = pgData[i][imgIdx];
    const newUrl = fixUrl(oldUrl);
    if (newUrl !== oldUrl){
      pgSheet.getRange(i + 1, imgIdx + 1).setValue(newUrl);
      fixedCount++;
    }
  }

  // Purchases.receipt_img / bill_img
  const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const pData = pSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const receiptIdx = pHeaders.indexOf('receipt_img');
  const billIdx = pHeaders.indexOf('bill_img');
  for (let i = 1; i < pData.length; i++){
    if (receiptIdx !== -1){
      const oldR = pData[i][receiptIdx];
      const newR = fixUrl(oldR);
      if (newR !== oldR){ pSheet.getRange(i + 1, receiptIdx + 1).setValue(newR); fixedCount++; }
    }
    if (billIdx !== -1){
      const oldB = pData[i][billIdx];
      const newB = fixUrl(oldB);
      if (newB !== oldB){ pSheet.getRange(i + 1, billIdx + 1).setValue(newB); fixedCount++; }
    }
  }

  Logger.log('แก้ไข URL รูปแล้ว: ' + fixedCount + ' รูป');
}

/**
 * ============================================================
 * ฟีเจอร์ QR Code — เพิ่มต่อท้ายไฟล์ Code.gs
 * ใช้สำหรับสร้างลิงก์ QR ที่เปิดตรงไปหน้ารายละเอียดสินค้า
 * ============================================================
 */
function getWebAppUrl(){
  return ScriptApp.getService().getUrl();
}


/**
 * อ่านชีต Purchases แค่ครั้งเดียว สรุป "รายการล่าสุดของแต่ละร้าน ต่อสินค้าแต่ละชิ้น"
 * ใช้ร่วมกันโดย getStoreRanking() และ getSwitchStoreSuggestions()
 */
function buildLatestByProductStore_(){
  const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const lastRow = pSheet.getLastRow();
  const latestByProductStore = {};
  const storeStats = {};
  if (lastRow <= 1) return { latestByProductStore, storeStats };

  const headers = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];
  const data = pSheet.getRange(2, 1, lastRow - 1, pSheet.getLastColumn()).getValues();
  const statusIdx = headers.indexOf('status');

  data.forEach(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    if (statusIdx !== -1 && obj.status === 'cancelled') return;

    const derived = calcPurchaseDerived(obj);
    const gid = obj.product_group_id, sid = obj.store_id;

    if (!storeStats[sid]) storeStats[sid] = { purchase_count: 0, total_spend: 0, product_ids: new Set() };
    storeStats[sid].purchase_count++;
    storeStats[sid].total_spend += derived.net_cost;
    storeStats[sid].product_ids.add(gid);

    if (!latestByProductStore[gid]) latestByProductStore[gid] = {};
    const existing = latestByProductStore[gid][sid];
    if (!existing || new Date(obj.date) > new Date(existing.date)){
      latestByProductStore[gid][sid] = {
        date: obj.date, cost_per_sell_unit: derived.cost_per_unit, unit_sell: obj.unit_sell
      };
    }
  });

  return { latestByProductStore, storeStats };
}


/**
 * ============================================================
 * 🏪 ร้านแนะนำ + 📉 วิเคราะห์แนวโน้ม/ช่วงเวลาซื้อถูกสุด
 * เพิ่มต่อท้ายไฟล์ Code.gs
 * ต่อยอดจากข้อมูลที่มีอยู่แล้วทั้งหมด ไม่ต้องเพิ่มคอลัมน์/migrate ใดๆ
 * ============================================================
 */

function getStoreRanking(){
  const stores = getAllStores();
  const { latestByProductStore, storeStats } = buildLatestByProductStore_();

  const stats = {};
  stores.forEach(s => {
    const st = storeStats[s.id];
    stats[s.id] = {
      store_id: s.id, store_name: s.name,
      purchase_count: st ? st.purchase_count : 0,
      total_spend: st ? st.total_spend : 0,
      product_ids: st ? st.product_ids : new Set(),
      cheapest_count: 0
    };
  });

  Object.values(latestByProductStore).forEach(storeMap => {
    const entries = Object.entries(storeMap);
    if (entries.length === 0) return;
    const [cheapestStoreId] = entries.reduce((min, cur) => cur[1].cost_per_sell_unit < min[1].cost_per_sell_unit ? cur : min, entries[0]);
    if (stats[cheapestStoreId]) stats[cheapestStoreId].cheapest_count++;
  });

  return Object.values(stats)
    .filter(s => s.purchase_count > 0)
    .map(s => ({
      store_id: s.store_id, store_name: s.store_name,
      purchase_count: s.purchase_count,
      total_spend: Math.round(s.total_spend * 100) / 100,
      product_variety: s.product_ids.size,
      cheapest_count: s.cheapest_count
    }))
    .sort((a, b) => b.cheapest_count - a.cheapest_count || b.purchase_count - a.purchase_count);
}


/** วิเคราะห์ช่วงเวลา (เดือน) ที่ราคาซื้อถูก/แพงที่สุด ของกลุ่มสินค้าหนึ่ง + แนวโน้มล่าสุด */
function getSeasonalPriceAnalysis(productGroupId){
  const purchases = getPurchasesByProductGroup(productGroupId);
  if (purchases.length === 0) return null;

  const monthNames = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const monthStats = {};
  purchases.forEach(p => {
    const m = new Date(p.date).getMonth() + 1;
    if (!monthStats[m]) monthStats[m] = { sum: 0, count: 0 };
    monthStats[m].sum += p.cost_per_unit;
    monthStats[m].count++;
  });
  const monthAvgs = Object.keys(monthStats).map(m => ({
    month: Number(m), month_name: monthNames[Number(m) - 1], avg: monthStats[m].sum / monthStats[m].count
  }));

  const allCosts = purchases.map(p => p.cost_per_unit);
  const avgCost = allCosts.reduce((s, c) => s + c, 0) / allCosts.length;

  let cheapestMonth = null, priciestMonth = null;
  if (monthAvgs.length >= 2){
    const sortedByAvg = monthAvgs.slice().sort((a, b) => a.avg - b.avg);
    cheapestMonth = { ...sortedByAvg[0], avg: Math.round(sortedByAvg[0].avg * 100) / 100 };
    priciestMonth = { ...sortedByAvg[sortedByAvg.length - 1], avg: Math.round(sortedByAvg[sortedByAvg.length - 1].avg * 100) / 100 };
  }

  let trend = null;
  const sorted = purchases.slice().sort((a, b) => new Date(a.date) - new Date(b.date));
  if (sorted.length >= 4){
    const half = Math.floor(sorted.length / 2);
    const olderAvg = sorted.slice(0, half).reduce((s, p) => s + p.cost_per_unit, 0) / half;
    const newerAvg = sorted.slice(half).reduce((s, p) => s + p.cost_per_unit, 0) / (sorted.length - half);
    const pctChange = olderAvg > 0 ? ((newerAvg - olderAvg) / olderAvg * 100) : 0;
    trend = {
      direction: pctChange > 2 ? 'up' : (pctChange < -2 ? 'down' : 'flat'),
      percent: Math.round(Math.abs(pctChange) * 10) / 10
    };
  }

  return {
    purchase_count: purchases.length,
    avg_cost: Math.round(avgCost * 100) / 100,
    cheapest_month: cheapestMonth,
    priciest_month: priciestMonth,
    trend: trend
  };
}

/**
 * ============================================================
 * 💰 เงินที่ประหยัดได้ + 🔁 วิเคราะห์ความถี่การซื้อ (แจ้งเตือนควรสั่งเพิ่ม)
 * เพิ่มต่อท้ายไฟล์ Code.gs — ต่อยอดจากข้อมูลเดิมทั้งหมด ไม่ต้อง migrate
 * ============================================================
 */

/** สรุปเดือนนี้: ใช้จ่ายรวม + ประหยัดไปเท่าไรจากการซื้อต่ำกว่าราคาเฉลี่ยของสินค้านั้นๆ */

 function getMonthlySavings(year, month){
  const now = new Date();
  year = year || now.getFullYear();
  month = month || (now.getMonth() + 1);


  const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const lastRow = pSheet.getLastRow();
  if (lastRow <= 1) return { year, month, total_spend: 0, purchase_count: 0, store_count: 0, total_saved: 0 };

  const headers = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];
  const data = pSheet.getRange(2, 1, lastRow - 1, pSheet.getLastColumn()).getValues();
  const statusIdx = headers.indexOf('status');

  const productCosts = {}; // product_group_id -> {sum, count} ต้นทุน/หน่วยขาย all-time
  const allPurchases = [];

  data.forEach(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    if (statusIdx !== -1 && obj.status === 'cancelled') return;
    const derived = calcPurchaseDerived(obj);
    const full = { ...obj, ...derived };
    allPurchases.push(full);
    if (!productCosts[obj.product_group_id]) productCosts[obj.product_group_id] = { sum: 0, count: 0 };
    productCosts[obj.product_group_id].sum += derived.cost_per_unit;
    productCosts[obj.product_group_id].count++;
  });

  let totalSpend = 0, totalSaved = 0, purchaseCount = 0;
  const storeSet = new Set();

  allPurchases.forEach(p => {
    const d = new Date(p.date);
    if (d.getFullYear() !== year || (d.getMonth() + 1) !== month) return;
    purchaseCount++;
    storeSet.add(p.store_id);
    totalSpend += p.net_cost;

    const pc = productCosts[p.product_group_id];
    const avgCost = pc.sum / pc.count;
    if (p.cost_per_unit < avgCost){
      totalSaved += (avgCost - p.cost_per_unit) * p.qty_sell_total;
    }
  });

  return {
    year: year, month: month,
    total_spend: Math.round(totalSpend * 100) / 100,
    purchase_count: purchaseCount,
    store_count: storeSet.size,
    total_saved: Math.round(totalSaved * 100) / 100
  };
}

/** ความถี่การซื้อของสินค้าหนึ่ง + เช็คว่าถึงเวลาควรสั่งเพิ่มหรือยัง */
function getPurchaseFrequencyInfo(productGroupId){
  const purchases = getPurchasesByProductGroup(productGroupId);
  if (purchases.length === 0) return null;

  const latest = purchases[0]; // เรียงใหม่ → เก่าอยู่แล้ว
  const daysSinceLast = Math.floor((new Date() - new Date(latest.date)) / 86400000);

  let avgIntervalDays = null;
  if (purchases.length >= 2){
    const datesAsc = purchases.map(p => new Date(p.date)).sort((a, b) => a - b);
    let totalDays = 0;
    for (let i = 1; i < datesAsc.length; i++) totalDays += (datesAsc[i] - datesAsc[i - 1]) / 86400000;
    avgIntervalDays = Math.round(totalDays / (datesAsc.length - 1));
  }

  const shouldReorder = avgIntervalDays !== null && daysSinceLast >= avgIntervalDays;

  return {
    purchase_count: purchases.length,
    last_purchase_date: latest.date,
    days_since_last_purchase: daysSinceLast,
    avg_interval_days: avgIntervalDays,
    should_reorder: shouldReorder,
    days_overdue: shouldReorder ? (daysSinceLast - avgIntervalDays) : 0
  };
}

/**
 * ============================================================
 * ⚠️ REPLACE — แทนที่ getProductGroupCards() เดิม ด้วยเวอร์ชันนี้
 * (เพิ่ม days_since_last_purchase / should_reorder สำหรับติดป้าย "ควรซื้อเพิ่ม" บนการ์ด)
 * ============================================================
 */
function getProductGroupCards(categoryId) {
  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const lastRow = pgSheet.getLastRow();
  if (lastRow <= 1) return [];

  const headers = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
  const data = pgSheet.getRange(2, 1, lastRow - 1, pgSheet.getLastColumn()).getValues();

  return data
    .filter(row => row[headers.indexOf('category_id')] === categoryId && row[headers.indexOf('active')] !== false)
    .map(row => {
      const pg = {};
      headers.forEach((h, i) => pg[h] = row[i]);

      const purchases = getPurchasesByProductGroup(pg.id);
      const storeIds = new Set(purchases.map(p => p.store_id));
      const latest = purchases[0];

      const costPerSellUnit = latest ? latest.cost_per_unit : 0;
      const costPerBuyUnit = latest ? Math.round((latest.net_cost / latest.qty_buy) * 100) / 100 : 0;
      const unitBuy = latest ? latest.unit_buy : '';
      const unitSell = latest ? latest.unit_sell : '';
      const ratio = latest ? latest.ratio : 0;

      const sellPricePerSellUnit = pg.current_sell_price || 0;
      const sellPricePerBuyUnit = ratio ? Math.round(sellPricePerSellUnit * ratio * 100) / 100 : 0;

      let purchaseFrequencyDays = null;
      if (purchases.length >= 2) {
        const datesAsc = purchases.map(p => new Date(p.date)).sort((a, b) => a - b);
        let totalDays = 0;
        for (let i = 1; i < datesAsc.length; i++) {
          totalDays += (datesAsc[i] - datesAsc[i - 1]) / (1000 * 60 * 60 * 24);
        }
        purchaseFrequencyDays = Math.round(totalDays / (datesAsc.length - 1));
      }

      const daysSinceLastPurchase = latest ? Math.floor((new Date() - new Date(latest.date)) / 86400000) : null;
      const shouldReorder = purchaseFrequencyDays !== null && daysSinceLastPurchase !== null && daysSinceLastPurchase >= purchaseFrequencyDays;

      return {
        id: pg.id, name: pg.name, image_url: pg.image_url,
        last_purchase_date: latest ? latest.date : null,
        cost_per_buy_unit: costPerBuyUnit, unit_buy: unitBuy,
        cost_per_sell_unit: costPerSellUnit, unit_sell: unitSell,
        sell_price_per_buy_unit: sellPricePerBuyUnit,
        sell_price_per_sell_unit: sellPricePerSellUnit,
        store_count: storeIds.size,
        purchase_count: purchases.length,
        purchase_frequency_days: purchaseFrequencyDays,
        days_since_last_purchase: daysSinceLastPurchase,
        should_reorder: shouldReorder
      };
    });
}

/**
 * ============================================================
 * ⚠️ REPLACE — แทนที่ getProductGroupDetail() เดิม ด้วยเวอร์ชันนี้
 * (เพิ่ม frequency object สำหรับแสดงบนหน้ารายละเอียดสินค้า)
 * ============================================================
 */
function getProductGroupDetail(productGroupId) {
  const pricing = getPricingInfo(productGroupId);
  const purchases = getPurchasesByProductGroup(productGroupId);
  const latest = purchases[0];

  const stores = getStoreListByProductGroup(productGroupId);

  let cheapestId = null, expensiveId = null, latestId = null;
  if (stores.length > 0) {
    cheapestId = stores.reduce((min, s) => s.cost_per_sell_unit < min.cost_per_sell_unit ? s : min, stores[0]).store_id;
    expensiveId = stores.reduce((max, s) => s.cost_per_sell_unit > max.cost_per_sell_unit ? s : max, stores[0]).store_id;
    latestId = stores.reduce((lat, s) => new Date(s.last_purchase_date) > new Date(lat.last_purchase_date) ? s : lat, stores[0]).store_id;
  }

  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
  const pgRow = pgSheet.getDataRange().getValues().find(r => r[0] === productGroupId);
  if (!pgRow) throw new Error('ไม่พบกลุ่มสินค้านี้');
  const urlIdx = pgHeaders.indexOf('product_url');

  const frequency = getPurchaseFrequencyInfo(productGroupId);

  return {
    ...pricing,
    name: pgRow[2], image_url: pgRow[3], product_url: urlIdx !== -1 ? (pgRow[urlIdx] || '') : '',
    cost_per_buy_unit: latest ? Math.round((latest.net_cost / latest.qty_buy) * 100) / 100 : 0,
    unit_buy: latest ? latest.unit_buy : '',
    unit_sell: latest ? latest.unit_sell : '',
    frequency: frequency,
    stores: stores.map(s => ({
      ...s,
      is_cheapest: s.store_id === cheapestId,
      is_most_expensive: s.store_id === expensiveId && expensiveId !== cheapestId,
      is_latest: s.store_id === latestId
    }))
  };
}


/**
 * ============================================================
 * 💡 สินค้าที่ควรเปลี่ยนร้าน + 📊 Dashboard รายเดือน (รวม #3, #5, #8, #10)
 * เพิ่มต่อท้ายไฟล์ Code.gs — ต่อยอดจากข้อมูลเดิมทั้งหมด ไม่ต้อง migrate
 * ============================================================
 */

/** สินค้าที่ซื้อล่าสุดจากร้านที่ไม่ใช่ร้านถูกที่สุด — แนะนำให้เปลี่ยนร้าน */
function getSwitchStoreSuggestions(){
  const { latestByProductStore } = buildLatestByProductStore_();

  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgLastRow = pgSheet.getLastRow();
  const productNames = {};
  if (pgLastRow > 1){
    const pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
    const pgData = pgSheet.getRange(2, 1, pgLastRow - 1, pgSheet.getLastColumn()).getValues();
    const idIdx = pgHeaders.indexOf('id'), nameIdx = pgHeaders.indexOf('name');
    pgData.forEach(r => productNames[r[idIdx]] = r[nameIdx]);
  }

  const storesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  const storesData = storesSheet.getRange(2, 1, Math.max(storesSheet.getLastRow() - 1, 0), 2).getValues();
  const storeNames = Object.fromEntries(storesData.map(r => [r[0], r[1]]));

  const suggestions = [];
  Object.keys(latestByProductStore).forEach(gid => {
    const entries = Object.entries(latestByProductStore[gid]);
    if (entries.length < 2) return;

    const [latestStoreId, latestInfo] = entries.reduce((lat, cur) => new Date(cur[1].date) > new Date(lat[1].date) ? cur : lat, entries[0]);
    const [cheapestStoreId, cheapestInfo] = entries.reduce((min, cur) => cur[1].cost_per_sell_unit < min[1].cost_per_sell_unit ? cur : min, entries[0]);
    if (cheapestStoreId === latestStoreId) return;

    const savingsPerUnit = Math.round((latestInfo.cost_per_sell_unit - cheapestInfo.cost_per_sell_unit) * 100) / 100;
    if (savingsPerUnit <= 0) return;

    suggestions.push({
      product_group_id: gid, product_name: productNames[gid] || gid, unit_sell: latestInfo.unit_sell,
      current_store_name: storeNames[latestStoreId] || latestStoreId, current_price: latestInfo.cost_per_sell_unit,
      cheapest_store_name: storeNames[cheapestStoreId] || cheapestStoreId, cheapest_price: cheapestInfo.cost_per_sell_unit,
      savings_per_unit: savingsPerUnit
    });
  });

  return suggestions.sort((a, b) => b.savings_per_unit - a.savings_per_unit);
}

// เพิ่มฟังก์ชันดึงข้อมูล (อ่านชีตแค่ครั้งเดียว กันช้า)
function getMasterDataForOCR(){
  const categories = getCategories(); // [{id, name}] มีอยู่แล้ว

  const pgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductGroups');
  const pgLastRow = pgSheet.getLastRow();
  if (pgLastRow <= 1) return { categories, productGroups: [] };

  const pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
  const pgData = pgSheet.getRange(2, 1, pgLastRow - 1, pgSheet.getLastColumn()).getValues();
  const activeIdx = pgHeaders.indexOf('active');
  const idIdx = pgHeaders.indexOf('id');
  const catIdx = pgHeaders.indexOf('category_id');
  const nameIdx = pgHeaders.indexOf('name');

  // อ่านชีต Purchases แค่ครั้งเดียว หาแถวล่าสุดของแต่ละกลุ่มสินค้า (กันปัญหาช้าแบบ N+1)
  const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const pLastRow = pSheet.getLastRow();
  const latestByGroup = {};
  if (pLastRow > 1){
    const pHeaders = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];
    const pData = pSheet.getRange(2, 1, pLastRow - 1, pSheet.getLastColumn()).getValues();
    const statusIdx = pHeaders.indexOf('status');
    const gidIdx = pHeaders.indexOf('product_group_id');
    const dateIdx = pHeaders.indexOf('date');
    const ubIdx = pHeaders.indexOf('unit_buy');
    const usIdx = pHeaders.indexOf('unit_sell');
    const rIdx = pHeaders.indexOf('ratio');

    pData.forEach(row => {
      if (statusIdx !== -1 && row[statusIdx] === 'cancelled') return;
      const gid = row[gidIdx];
      const existing = latestByGroup[gid];
      if (!existing || new Date(row[dateIdx]) > new Date(existing.date)){
        latestByGroup[gid] = { date: row[dateIdx], unit_buy: row[ubIdx], unit_sell: row[usIdx], ratio: row[rIdx] };
      }
    });
  }

  const productGroups = pgData
    .filter(row => row[activeIdx] !== false)
    .map(row => {
      const pgId = row[idIdx];
      const latest = latestByGroup[pgId];
      return {
        id: pgId, name: row[nameIdx], category_id: row[catIdx],
        has_history: !!latest,
        unit_buy: latest ? latest.unit_buy : '',
        unit_sell: latest ? latest.unit_sell : '',
        ratio: latest ? latest.ratio : ''
      };
    });

  return { categories, productGroups };
}










/** Dashboard รายเดือน: สรุปยอด + ดัชนีราคา + สินค้าที่ควรซื้อเพิ่ม + insight อัตโนมัติ */
function getDashboardData(year, month){
  const now = new Date();
  year = year || now.getFullYear();
  month = month || (now.getMonth() + 1);
  const savings = getMonthlySavings(year, month);
  // ลบบรรทัด "const year = now.getFullYear(), month = now.getMonth() + 1;" เดิมที่อยู่ถัดไปทิ้ง (ซ้ำกับด้านบน)
  // ที่เหลือทั้งหมด (priceDropCount, indexChangePercent, reorderList, switchSuggestions, storeRanking) ใช้โค้ดเดิมได้เลย ไม่ต้องแก้
  // จำนวนสินค้าที่ราคาขายเปลี่ยนแปลงเดือนนี้
  let priceDropCount = 0, priceRiseCount = 0;
  const phSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PriceHistory');
  const phLastRow = phSheet.getLastRow();
  if (phLastRow > 1){
    const phHeaders = phSheet.getRange(1, 1, 1, phSheet.getLastColumn()).getValues()[0];
    const phData = phSheet.getRange(2, 1, phLastRow - 1, phSheet.getLastColumn()).getValues();
    const dateIdx = phHeaders.indexOf('date'), oldIdx = phHeaders.indexOf('old_price'), newIdx = phHeaders.indexOf('new_price');
    phData.forEach(row => {
      const d = new Date(row[dateIdx]);
      if (d.getFullYear() !== year || (d.getMonth() + 1) !== month) return;
      if (row[newIdx] < row[oldIdx]) priceDropCount++;
      else if (row[newIdx] > row[oldIdx]) priceRiseCount++;
    });
  }

  // ดัชนีราคารวม: ต้นทุนเฉลี่ยถ่วงน้ำหนัก 3 เดือนล่าสุด เทียบ 3 เดือนก่อนหน้า (ทุกสินค้ารวมกัน)
  const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Purchases');
  const pLastRow = pSheet.getLastRow();
  let indexChangePercent = null;
  if (pLastRow > 1){
    const pHeaders = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];
    const pData = pSheet.getRange(2, 1, pLastRow - 1, pSheet.getLastColumn()).getValues();
    const statusIdx = pHeaders.indexOf('status');

    const cutoff3 = new Date(); cutoff3.setDate(cutoff3.getDate() - 90);
    const cutoff6 = new Date(); cutoff6.setDate(cutoff6.getDate() - 180);

    let recentCostQty = 0, recentQty = 0, priorCostQty = 0, priorQty = 0;
    pData.forEach(row => {
      const obj = {};
      pHeaders.forEach((h, i) => obj[h] = row[i]);
      if (statusIdx !== -1 && obj.status === 'cancelled') return;
      const derived = calcPurchaseDerived(obj);
      const d = new Date(obj.date);
      if (d >= cutoff3){
        recentCostQty += derived.cost_per_unit * derived.qty_sell_total;
        recentQty += derived.qty_sell_total;
      } else if (d >= cutoff6){
        priorCostQty += derived.cost_per_unit * derived.qty_sell_total;
        priorQty += derived.qty_sell_total;
      }
    });

    if (recentQty > 0 && priorQty > 0){
      const recentAvg = recentCostQty / recentQty;
      const priorAvg = priorCostQty / priorQty;
      if (priorAvg > 0) indexChangePercent = Math.round(((recentAvg - priorAvg) / priorAvg) * 1000) / 10;
    }
  }

  // สินค้าที่ถึงรอบควรซื้อเพิ่ม
  //const categories = getCategories();
 // const reorderList = [];
  const reorderList = getProductGroupCardsBatch(null).filter(c => c.should_reorder);


  const switchSuggestions = getSwitchStoreSuggestions();
  const storeRanking = getStoreRanking();

  // สรุปอัตโนมัติแบบ rule-based (ไม่ใช่ AI จริง แต่สร้างจากข้อมูลที่คำนวณไว้แล้ว)
  const insights = [];
  if (indexChangePercent !== null){
    if (indexChangePercent <= -2) insights.push(`📉 ต้นทุนโดยรวมลดลง ${Math.abs(indexChangePercent)}% เทียบกับ 3 เดือนก่อน`);
    else if (indexChangePercent >= 2) insights.push(`📈 ต้นทุนโดยรวมเพิ่มขึ้น ${indexChangePercent}% เทียบกับ 3 เดือนก่อน`);
    else insights.push(`➡️ ต้นทุนโดยรวมค่อนข้างทรงตัวเทียบกับ 3 เดือนก่อน`);
  }
  if (storeRanking.length > 0){
    insights.push(`🏆 ${storeRanking[0].store_name} เป็นร้านที่คุ้มที่สุดตอนนี้ (ถูกที่สุด ${storeRanking[0].cheapest_count} รายการ)`);
  }
  if (reorderList.length > 0){
    insights.push(`⏰ มี ${reorderList.length} รายการที่ถึงรอบควรซื้อเพิ่มแล้ว`);
  }
  if (switchSuggestions.length > 0){
    const top = switchSuggestions[0];
    insights.push(`💡 เปลี่ยนร้านซื้อ "${top.product_name}" เป็น ${top.cheapest_store_name} ประหยัดได้ ฿${top.savings_per_unit}/${top.unit_sell}`);
  }
  if (priceDropCount > 0) insights.push(`✅ ปรับราคาขายลดลงไปแล้ว ${priceDropCount} รายการเดือนนี้`);
  if (insights.length === 0) insights.push('ยังไม่มีข้อมูลพอสรุป ลองบันทึกการซื้อเพิ่มแล้วกลับมาดูใหม่');

  return {
    ...savings,
    price_drop_count: priceDropCount,
    price_rise_count: priceRiseCount,
    index_change_percent: indexChangePercent,
    reorder_count: reorderList.length,
    switch_suggestion_count: switchSuggestions.length,
    insights: insights
  };
}
// ---End Code.gs ---

## Index.html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <?!= include('Stylesheet'); ?>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jsQR/1.4.0/jsQR.js"></script>
 
  
</head>
<body>
  <div id="app">

    <div class="topbar" id="topbar"></div>

    <div class="content" id="content">
      <div class="spinner"></div>
    </div>

    <div class="bottom-nav" id="bottomNav"></div>

    <div class="toast" id="toast"></div>
  </div>
<script>
  window.__initialPage = "<?= initialPage ?>";
  window.__initialId = "<?= initialId ?>";
  //alert('page=[' + window.__initialPage + '] pgId=[' + window.__initialId + ']');
</script>


  <?!= include('JavaScript'); ?>
</body>
</html>
  

## AccessDenied.html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <?!= include('Stylesheet'); ?>
</head>
<body>
  <div style="max-width:420px; margin:80px auto; text-align:center; padding:0 24px;">
    <div style="font-size:52px; margin-bottom:16px;">🔒</div>
    <h1 style="font-size:22px; margin-bottom:10px;">ยังไม่มีสิทธิ์เข้าถึง</h1>
    <p style="color:var(--ink-soft); font-size:15px; line-height:1.6;">
      บัญชี Google ของคุณยังไม่ได้รับอนุญาตให้เข้าใช้งานระบบนี้<br>
      กรุณาติดต่อผู้ดูแลระบบเพื่อขอสิทธิ์เข้าถึง
    </p>
  </div>
</body>
</html>

## Stylesheet.html
<style>
  
  @import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@500;700;800&family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap');

  /* ==========================================================
     TOKENS — โทนร้านเครื่องเขียน: ครีมอ่อน + ส้มโคโด่ + มิ้นต์ + เหลืองแดด
     ========================================================== */
  :root{
    --bg:        #FFFBF5;
    --card:      #FFFFFF;
    --orange:    #FF6B4A;
    --orange-dk: #E5502F;
    --mint:      #00BFA6;
    --mint-dk:   #049682;
    --yellow:    #FFC93C;
    --ink:       #3A2E28;
    --ink-soft:  #8C7B6F;
    --line:      #F0DDC8;
    --line-dash: #E7C9A8;
    --danger:    #E5484D;
    --radius-lg: 20px;
    --radius-md: 14px;
    --shadow:    0 6px 18px rgba(58,46,40,0.08);
  }

  @font-face{ /* fallback handled by @import at top of file */ }

  *{ box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
  html,body{
    margin:0; padding:0; background:var(--bg); color:var(--ink);
    font-family:'IBM Plex Sans Thai','Baloo 2',sans-serif;
    -webkit-font-smoothing:antialiased;
  }
  body{ padding-bottom:78px; } /* กันชนแถบล่าง */

  .num{ font-family:'Baloo 2','IBM Plex Sans Thai',sans-serif; font-variant-numeric:tabular-nums; font-feature-settings:'tnum' 1; }

  h1,h2,h3{ font-family:'Baloo 2','IBM Plex Sans Thai',sans-serif; margin:0; color:var(--ink); }

  button{ font-family:inherit; }

  /* ==========================================================
     APP SHELL
     ========================================================== */
  #app{ max-width:520px; margin:0 auto; min-height:100vh; position:relative; }

  .topbar{
    position:sticky; top:0; z-index:20; background:var(--bg);
    padding:16px 16px 10px; border-bottom:1px solid var(--line);
  }
  .topbar-row{ display:flex; align-items:center; gap:10px; }
  .back-btn{
    width:40px; height:40px; border-radius:50%; border:none; background:var(--card);
    box-shadow:var(--shadow); font-size:18px; color:var(--ink); display:flex;
    align-items:center; justify-content:center; flex-shrink:0;
  }
  .topbar h1{ font-size:20px; font-weight:800; flex:1; }

  .searchbox{
    margin-top:12px; display:flex; align-items:center; gap:8px;
    background:var(--card); border:1px solid var(--line); border-radius:999px;
    padding:10px 16px; box-shadow:var(--shadow);
  }
  .searchbox input{
    border:none; outline:none; flex:1; font-size:15px; background:transparent; color:var(--ink);
    font-family:'IBM Plex Sans Thai',sans-serif;
  }
  .searchbox svg{ flex-shrink:0; opacity:.5; }

  .tabs{ display:flex; gap:8px; overflow-x:auto; padding:12px 16px 4px; scrollbar-width:none; }
  .tabs::-webkit-scrollbar{ display:none; }
  .tab{
    flex-shrink:0; padding:8px 16px; border-radius:999px; border:1.5px solid var(--line);
    background:var(--card); color:var(--ink-soft); font-size:14px; font-weight:600;
    white-space:nowrap;
  }
  .tab.active{ background:var(--orange); border-color:var(--orange); color:#fff; }

  .content{ padding:14px 16px 24px; }

  /* ==========================================================
     PRICE-TAG CARD — ลายเซ็นของทั้งแอป
     รูเจาะที่มุมซ้ายบน + ขอบปะติดแบบตัดเก็บ + เส้นปรุฉีก
     ========================================================== */
  .tag-card{
    position:relative; background:var(--card); border-radius:var(--radius-lg);
    box-shadow:var(--shadow); overflow:hidden; border:1px solid var(--line);
    margin-bottom:14px;
  }
  .tag-hole{
    position:absolute; top:14px; left:14px; width:14px; height:14px; border-radius:50%;
    background:var(--bg); border:2px solid var(--line-dash); z-index:2;
  }
  .tag-media{
    height:120px; background:linear-gradient(135deg,#FFF3E9,#FFE7DB);
    display:flex; align-items:center; justify-content:center; font-size:38px;
  }
  .tag-media img{ width:100%; height:100%; object-fit:cover; }
  .tag-body{ padding:14px 16px 16px 34px; }
  .tag-name{ font-size:16px; font-weight:700; margin-bottom:6px; line-height:1.3; }
  .tag-perforation{
    border-top:2px dashed var(--line-dash); margin:0 16px;
  }
  .tag-price-row{
    display:flex; align-items:baseline; justify-content:space-between;
    padding:10px 16px 12px 34px;
  }
  .tag-price{ font-size:22px; font-weight:800; color:var(--orange-dk); }
  .tag-price small{ font-size:12px; font-weight:600; color:var(--ink-soft); margin-left:2px; }
  .tag-meta{ font-size:12.5px; color:var(--ink-soft); display:flex; gap:10px; }

  .badge{
    display:inline-flex; align-items:center; gap:4px; font-size:11.5px; font-weight:700;
    padding:3px 9px; border-radius:999px;
  }
  .badge-mint{ background:rgba(0,191,166,0.14); color:var(--mint-dk); }
  .badge-yellow{ background:rgba(255,201,60,0.22); color:#8A6300; }
  .badge-green{ background:rgba(0,191,166,0.14); color:var(--mint-dk); }
  .badge-amber{ background:rgba(255,201,60,0.22); color:#8A6300; }
  .badge-red{ background:rgba(229,72,77,0.12); color:var(--danger); }

  .grid{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }

  /* ==========================================================
     PRODUCT CARD (หน้าแรก) — แนวนอน เน้นข้อมูลการซื้อ
     ========================================================== */
  .product-card{
    position:relative; display:flex; gap:12px; background:var(--card); border:1px solid var(--line);
    border-radius:var(--radius-lg); box-shadow:var(--shadow); padding:14px; margin-bottom:12px; cursor:pointer;
  }
  .product-card-media{
    position:relative; width:84px; height:84px; flex-shrink:0; border-radius:var(--radius-md); overflow:hidden;
    background:linear-gradient(135deg,#FFF3E9,#FFE7DB); display:flex; align-items:center; justify-content:center;
  }
  .product-card-media img{ width:100%; height:100%; object-fit:cover; }
  .product-card-emoji{ font-size:30px; }
  .pc-upload-btn{
    position:absolute; bottom:4px; right:4px; width:26px; height:26px; background:rgba(255,255,255,0.92);
    box-shadow:var(--shadow); font-size:12px;
  }
  .product-card-info{ flex:1; min-width:0; display:flex; flex-direction:column; gap:3px; justify-content:center; }
  .product-card-name{ font-weight:700; font-size:15px; line-height:1.25; }
  .product-card-date{ font-size:11.5px; color:var(--ink-soft); }
  .product-card-cost-row{ display:flex; gap:8px; margin-top:3px; flex-wrap:wrap; }
  .cost-chip{ display:flex; align-items:baseline; gap:2px; padding:4px 9px; border-radius:999px; }
  .cost-chip.cost-buy{ background:rgba(255,107,74,0.13); }
  .cost-chip.cost-buy .cost-value{ color:var(--orange-dk); font-weight:800; font-size:14.5px; }
  .cost-chip.cost-sell{ background:rgba(58,46,40,0.045); }
  .cost-chip.cost-sell .cost-value{ color:var(--ink); font-weight:700; font-size:13px; }
  .cost-chip .cost-unit{ font-size:10.5px; color:var(--ink-soft); }
  .product-card-sell-row{ font-size:10.5px; color:var(--ink-soft); opacity:.7; }
  .product-card-meta{ font-size:11.5px; color:var(--ink-soft); margin-top:2px; }

  /* generic simple list row card */
  .row-card{
    display:flex; align-items:center; gap:12px; background:var(--card);
    border:1px solid var(--line); border-radius:var(--radius-md); padding:14px;
    box-shadow:var(--shadow); margin-bottom:10px;
  }
  .row-card .avatar{
    width:44px; height:44px; border-radius:50%; background:var(--yellow);
    display:flex; align-items:center; justify-content:center; font-weight:800; color:#fff;
    flex-shrink:0; font-size:16px;
  }
  .row-main{ flex:1; min-width:0; }
  .row-title{ font-weight:700; font-size:15px; display:flex; align-items:center; gap:6px; }
  .row-sub{ font-size:12.5px; color:var(--ink-soft); margin-top:2px; }
  .row-price{ text-align:right; font-weight:800; font-size:16px; color:var(--orange-dk); }
  .row-arrow{ color:var(--ink-soft); flex-shrink:0; }
  .row-card.cancelled{ opacity:.55; }
  .row-card.cancelled .row-title{ text-decoration:line-through; text-decoration-color:var(--ink-soft); }
  .row-card.cancelled .row-price{ color:var(--ink-soft); text-decoration:line-through; }

  /* ==========================================================
     SECTIONS / DETAIL PANELS
     ========================================================== */
  .section-title{
    font-size:13px; font-weight:700; color:var(--ink-soft); text-transform:uppercase;
    letter-spacing:.04em; margin:22px 0 10px; display:flex; align-items:center; gap:6px;
  }

  .stat-tag{
    position:relative; background:var(--card); border:1px solid var(--line);
    border-radius:var(--radius-md); padding:14px 14px 14px 30px; box-shadow:var(--shadow);
  }
  .stat-tag .tag-hole{ top:10px; left:10px; width:10px; height:10px; }
  .stat-tag .stat-label{ font-size:12px; color:var(--ink-soft); margin-bottom:4px; }
  .stat-tag .stat-value{ font-size:19px; font-weight:800; }

  .status-banner{
    display:flex; align-items:center; gap:10px; padding:14px 16px; border-radius:var(--radius-md);
    margin-bottom:14px; font-weight:700; font-size:15px;
  }
  .status-green{ background:rgba(0,191,166,0.12); color:var(--mint-dk); }
  .status-yellow{ background:rgba(255,201,60,0.20); color:#8A6300; }
  .status-red{ background:rgba(229,72,77,0.10); color:var(--danger); }

  .info-block{
    background:var(--card); border:1px solid var(--line); border-radius:var(--radius-md);
    padding:16px; margin-bottom:14px; box-shadow:var(--shadow);
  }
  .info-block-title{ font-weight:700; font-size:14.5px; margin-bottom:10px; display:flex; align-items:center; gap:6px; }
  .info-row{ display:flex; justify-content:space-between; padding:7px 0; font-size:14px; border-bottom:1px dashed var(--line); }
  .info-row:last-child{ border-bottom:none; }
  .info-row .k{ color:var(--ink-soft); }
  .info-row .v{ font-weight:600; }
  .info-row.total{ margin-top:4px; padding-top:10px; border-top:1.5px solid var(--line); border-bottom:none; }
  .info-row.total .v{ color:var(--orange-dk); font-weight:800; font-size:16px; }

  /* ==========================================================
     FORMS
     ========================================================== */
  .field{ margin-bottom:14px; }
  .field label{ display:block; font-size:13px; font-weight:600; color:var(--ink-soft); margin-bottom:6px; }
  .field input, .field select, .field textarea{
    width:100%; padding:13px 14px; border-radius:var(--radius-md); border:1.5px solid var(--line);
    background:var(--card); font-size:15px; color:var(--ink); font-family:inherit; outline:none;
  }
  .field input:focus, .field select:focus, .field textarea:focus{ border-color:var(--orange); }
  .field-row{ display:flex; gap:10px; }
  .field-row .field{ flex:1; }

  .autocomplete{ position:relative; }
  .autocomplete-list{
    position:absolute; left:0; right:0; top:calc(100% + 4px); background:var(--card);
    border:1px solid var(--line); border-radius:var(--radius-md); box-shadow:var(--shadow);
    z-index:30; max-height:220px; overflow-y:auto;
  }
  .autocomplete-item{ padding:12px 14px; font-size:14.5px; border-bottom:1px solid var(--line); }
  .autocomplete-item:last-child{ border-bottom:none; }
  .autocomplete-item.create-new{ color:var(--orange-dk); font-weight:700; }

  .segmented{ display:flex; background:var(--line); border-radius:999px; padding:4px; margin-bottom:16px; }
  .segmented button{
    flex:1; border:none; background:transparent; padding:10px; border-radius:999px;
    font-size:14px; font-weight:700; color:var(--ink-soft);
  }
  .segmented button.active{ background:var(--card); color:var(--ink); box-shadow:var(--shadow); }

  .checkbox-row{ display:flex; align-items:center; gap:8px; font-size:14px; margin-bottom:16px; }
  .checkbox-row input{ width:18px; height:18px; }

  /* ==========================================================
     BUTTONS — ใหญ่ กดง่ายด้วยนิ้วโป้ง
     ========================================================== */
  .btn{
    display:flex; align-items:center; justify-content:center; gap:8px;
    width:100%; padding:16px; border-radius:999px; border:none; font-size:16px;
    font-weight:700; cursor:pointer;
  }
  .btn-primary{ background:var(--orange); color:#fff; box-shadow:0 8px 20px rgba(255,107,74,0.35); }
  .btn-primary:active{ background:var(--orange-dk); }
  .btn-ghost{ background:var(--card); color:var(--ink); border:1.5px solid var(--line); }
  .btn-mint{ background:var(--mint); color:#fff; box-shadow:0 8px 20px rgba(0,191,166,0.3); }

.bottom-nav{
  position:fixed; left:0; right:0; bottom:0; z-index:40;
  display:flex; background:var(--card); border-top:1px solid var(--line);
  box-shadow:0 -6px 18px rgba(58,46,40,0.08);
  padding:6px 4px calc(6px + env(safe-area-inset-bottom, 0px));
  max-width:520px; margin:0 auto;
}
.nav-tab{
  flex:1; display:flex; flex-direction:column; align-items:center; gap:2px;
  border:none; background:transparent; padding:8px 2px; border-radius:14px;
  color:var(--ink-soft); font-size:10.5px; font-weight:700;
}
.nav-tab .nav-icon{ font-size:20px; line-height:1; }
.nav-tab.active{ color:var(--orange-dk); }
.nav-tab.active .nav-icon{ transform:translateY(-1px); }


  /* ==========================================================
     MISC
     ========================================================== */
  .empty-state{ text-align:center; padding:60px 20px; color:var(--ink-soft); }
  .empty-state .icon{ font-size:44px; margin-bottom:10px; }
  .empty-state .title{ font-weight:700; color:var(--ink); margin-bottom:4px; }

  .spinner{
    width:28px; height:28px; border-radius:50%; border:3px solid var(--line);
    border-top-color:var(--orange); animation:spin .7s linear infinite; margin:40px auto;
  }
  @keyframes spin{ to{ transform:rotate(360deg); } }

  .toast{
    position:fixed; left:50%; bottom:100px; transform:translateX(-50%);
    background:var(--ink); color:#fff; padding:12px 20px; border-radius:999px;
    font-size:14px; font-weight:600; z-index:99; opacity:0; pointer-events:none;
    transition:opacity .25s, bottom .25s;
  }
  .toast.show{ opacity:1; bottom:110px; }

  .receipt-thumb{
    width:100%; border-radius:var(--radius-md); border:1px dashed var(--line-dash);
    margin-top:10px;
  }

  /* ==========================================================
     MASTER DATA — จัดการข้อมูล
     ========================================================== */
  .master-row{
    display:flex; align-items:center; gap:10px; background:var(--card);
    border:1px solid var(--line); border-radius:var(--radius-md); padding:12px 14px;
    box-shadow:var(--shadow); margin-bottom:8px;
  }
  .master-row.inactive{ opacity:.5; }
  .master-row.selected{ border-color:var(--orange); background:#FFF6F1; }
  .master-check{
    width:22px; height:22px; border-radius:50%; border:2px solid var(--line-dash);
    flex-shrink:0; display:flex; align-items:center; justify-content:center; font-size:13px; color:#fff;
  }
  .master-check.checked{ background:var(--orange); border-color:var(--orange); }
  .master-name{ flex:1; min-width:0; font-weight:600; font-size:14.5px; }
  .master-name input{
    width:100%; border:none; border-bottom:1.5px solid var(--orange); background:transparent;
    font-size:14.5px; font-weight:600; font-family:inherit; padding:2px 0; outline:none; color:var(--ink);
  }
  .master-usage{ font-size:11.5px; color:var(--ink-soft); flex-shrink:0; }
  .icon-btn{
    width:34px; height:34px; border-radius:50%; border:none; background:var(--bg);
    display:flex; align-items:center; justify-content:center; font-size:15px; flex-shrink:0; color:var(--ink);
  }
  .icon-btn.danger{ color:var(--danger); }
  .icon-btn:disabled{ opacity:.3; }

  .switch{
    position:relative; width:40px; height:24px; border-radius:999px; background:var(--line);
    flex-shrink:0; border:none;
  }
  .switch.on{ background:var(--mint); }
  .switch::after{
    content:''; position:absolute; top:3px; left:3px; width:18px; height:18px; border-radius:50%;
    background:#fff; transition:transform .15s;
  }
  .switch.on::after{ transform:translateX(16px); }

  .select-bar{
    position:fixed; left:50%; transform:translateX(-50%); bottom:84px; z-index:45;
    width:calc(100% - 32px); max-width:488px; background:var(--ink); color:#fff;
    border-radius:999px; padding:10px 10px 10px 20px; display:flex; align-items:center;
    justify-content:space-between; gap:10px; box-shadow:var(--shadow);
  }
  .select-bar button{
    background:var(--orange); color:#fff; border:none; border-radius:999px; padding:10px 18px;
    font-weight:700; font-size:14px;
  }
  .select-bar .cancel-select{ background:transparent; color:#fff; opacity:.7; padding:10px; }

  .merge-target-row{
    display:flex; align-items:center; gap:10px; background:var(--card); border:1.5px solid var(--line);
    border-radius:var(--radius-md); padding:14px; margin-bottom:8px;
  }
  .merge-target-row.picked{ border-color:var(--mint); background:rgba(0,191,166,0.06); }
  .radio-dot{
    width:20px; height:20px; border-radius:50%; border:2px solid var(--line-dash); flex-shrink:0;
    display:flex; align-items:center; justify-content:center;
  }
  .radio-dot.on::after{ content:''; width:10px; height:10px; border-radius:50%; background:var(--mint); }

  @media (prefers-reduced-motion: reduce){
    *{ animation-duration:0.001ms !important; transition-duration:0.001ms !important; }
  }
  

.modal-overlay{
  position:fixed; inset:0; z-index:100; background:rgba(58,46,40,0.45);
  display:flex; align-items:center; justify-content:center; padding:24px;
}
.modal-box{
  background:var(--card); border-radius:var(--radius-lg); padding:22px;
  max-width:360px; width:100%; box-shadow:0 20px 50px rgba(0,0,0,0.25);
}
.modal-message{ font-size:15px; line-height:1.6; color:var(--ink); margin-bottom:16px; }
.modal-input{
  width:100%; padding:12px 14px; border-radius:var(--radius-md); border:1.5px solid var(--line);
  background:var(--bg); font-size:15px; color:var(--ink); font-family:inherit; outline:none; margin-bottom:16px;
}
.modal-input:focus{ border-color:var(--orange); }
.modal-actions{ display:flex; gap:10px; }
.modal-actions .btn{ padding:12px; font-size:14.5px; }





  
.select-bar{
 position:fixed; 
 left:50%; 
 bottom:100px; 
 transform:translateX(-50%);
  //right: 84px;
 
  z-index: 99;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  background: var(--ink);
  color: #fff;
  border-radius: 18px;
  padding: 10px 12px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.25);
  max-width: calc(150vw - 96px);
  box-sizing: border-box;
  flex-wrap: wrap;          /* ถ้าล้นให้ตกบรรทัดแทนล้นออกจอ */
}

.select-bar > span{
  font-size: 12.5px;
  white-space: nowrap;
  flex-shrink: 0;
}

.select-bar > div{
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
}

.select-bar button{
  border: none;
  border-radius: 560px;
  font-size: 12px;
  font-weight: 700;
  padding: 8px 12px;
  white-space: nowrap;
  cursor: pointer;
}

.select-bar .cancel-select{
  background: rgba(255,255,255,0.15);
  color: #fff;
}

.select-bar button:not(.cancel-select){
  background: var(--orange);
  color: #fff;
}

.select-fab{
  position: fixed;
  right: 16px;
  bottom: 140px;
  z-index: 56;
  background: var(--ink);
  color: #fff;
  border: none;
  border-radius: 999px;
  padding: 12px 18px;
  font-size: 13.5px;
  font-weight: 700;
  box-shadow: 0 4px 14px rgba(0,0,0,0.28);
  display: flex;
  align-items: center;
  gap: 6px;
  cursor: pointer;
}
.select-fab .count-badge{
  background: var(--orange);
  color: #fff;
  border-radius: 999px;
  min-width: 20px;
  height: 20px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 11.5px;
  padding: 0 5px;
}
.select-sheet-overlay{
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.35);
  z-index: 57;
  display: flex;
  align-items: flex-end;
}
.select-sheet{
  width: 100%;
  background: var(--card);
  border-radius: 20px 20px 0 0;
  padding: 18px 16px 26px;
  box-shadow: 0 -4px 20px rgba(0,0,0,0.2);
}
.select-sheet-header{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 14px;
  font-weight: 700;
  font-size: 15px;
}
.select-sheet-header .close-x{
  border: none;
  background: none;
  font-size: 20px;
  color: var(--ink-soft);
}
.select-sheet-actions{
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.select-sheet-actions button{
  border: none;
  border-radius: var(--radius-md);
  padding: 13px;
  font-size: 14px;
  font-weight: 700;
  cursor: pointer;
}
.select-sheet-actions .primary-move{ background: var(--orange); color: #fff; }
.select-sheet-actions .primary-merge{ background: var(--mint-dk); color: #fff; }
.select-sheet-actions .cancel-select{ background: var(--bg); color: var(--ink-soft); border: 1.5px solid var(--line); }


</style>




## JavaScript.html
<script>
/* ============================================================
   JavaScript.gs
   ============================================================ */
/* ============================================================
   รูปภาพ — ย่อขนาดในเบราว์เซอร์ก่อนส่งขึ้น Drive (กันไฟล์ใหญ่/ช้า)
   ============================================================ */
function resizeImageFile(file, maxSize, quality){
  maxSize = maxSize || 1000; quality = quality || 0.82;
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        let w = img.width, h = img.height;
        if (w > h && w > maxSize){ h = Math.round(h * maxSize / w); w = maxSize; }
        else if (h >= w && h > maxSize){ w = Math.round(w * maxSize / h); h = maxSize; }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = () => reject(new Error('ไม่สามารถอ่านไฟล์รูปภาพนี้ได้'));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error('อ่านไฟล์ไม่สำเร็จ'));
    reader.readAsDataURL(file);
  });
}

async function uploadPurchaseImageAndGetUrl(purchaseId, imageType, file){
  const dataUrl = await resizeImageFile(file);
  const base64 = dataUrl.split(',')[1];
  const res = await gs('uploadPurchaseImage', purchaseId, imageType, base64, 'image/jpeg', purchaseId + '_' + imageType + '.jpg');
  return res.url;
}


async function handleProductImageUpload(productGroupId, file, onDone){
  if (!file) return;
  if (!file.type.startsWith('image/')){ toast('กรุณาเลือกไฟล์รูปภาพ'); return; }
  toast('กำลังอัปโหลดรูป...');
  try{
    const dataUrl = await resizeImageFile(file);
    const base64 = dataUrl.split(',')[1];
    const res = await gs('uploadProductGroupImage', productGroupId, base64, 'image/jpeg', productGroupId + '.jpg');
    toast('อัปโหลดรูปแล้ว');
    if (onDone) onDone(res.url);
  } catch(err){
    toast(err.message || 'อัปโหลดรูปไม่สำเร็จ');
  }
}


/* ============================================================
   ส่วนลดแบบยืดหยุ่น — หลายรายการ ชื่อไม่ตายตัว
   ============================================================ */
function discountRowEl(name, value, type){
  type = type === 'percent' ? 'percent' : 'amount';
  const row = el(`
    <div class="field-row discount-row" style="align-items:center;">
      <div class="field" style="flex:1.3; margin-bottom:0;"><input class="disc-name" placeholder="ชื่อส่วนลด เช่น สมาชิก, โปรโมชั่น" value="${name || ''}"></div>
      <div class="field" style="flex:0.8; margin-bottom:0;"><input class="disc-value" type="number" inputmode="decimal" placeholder="0" value="${value != null ? value : ''}"></div>
      <button type="button" class="disc-type-btn" style="flex-shrink:0; width:38px; height:44px; border-radius:12px; border:1.5px solid var(--line); background:var(--card); font-weight:800; font-size:13px; color:var(--ink-soft);">${type === 'percent' ? '%' : '฿'}</button>
      <button type="button" class="icon-btn danger disc-remove" style="flex-shrink:0;">✕</button>
    </div>`);
  row.dataset.type = type;
  const typeBtn = row.querySelector('.disc-type-btn');
  typeBtn.onclick = () => {
    const next = row.dataset.type === 'percent' ? 'amount' : 'percent';
    row.dataset.type = next;
    typeBtn.textContent = next === 'percent' ? '%' : '฿';
  };
  row.querySelector('.disc-remove').onclick = () => row.remove();
  return row;
}

function setupDiscountRows(listId, addBtnId, initialDiscounts){
  const list = document.getElementById(listId);
  const addBtn = document.getElementById(addBtnId);
  (initialDiscounts && initialDiscounts.length > 0 ? initialDiscounts : []).forEach(d => {
    list.appendChild(discountRowEl(d.name, d.amount, 'amount')); // รายการเก่าเก็บเป็นจำนวนเงินแล้ว แสดงเป็น ฿ เสมอ
  });
  addBtn.onclick = () => list.appendChild(discountRowEl('', '', 'amount'));
}


function collectDiscounts(listId, priceBuy){
  priceBuy = Number(priceBuy) || 0;
  return Array.from(document.querySelectorAll(`#${listId} .discount-row`))
    .map(row => {
      const type = row.dataset.type === 'percent' ? 'percent' : 'amount';
      const rawValue = parseFloat(row.querySelector('.disc-value').value || '0');
      const amount = type === 'percent' ? Math.round(priceBuy * rawValue / 100 * 100) / 100 : rawValue;
      return { name: row.querySelector('.disc-name').value.trim(), amount: amount };
    })
    .filter(d => d.amount > 0);
}

/* ============================================================
   Dialog กำหนดเอง — แทน confirm()/prompt() ของเบราว์เซอร์
   (ของเดิมโชว์ข้อความ "หน้าที่ฝังไว้ใน [url ยาวๆ] บอกว่า" ดูไม่เป็นมืออาชีพ)
   ============================================================ */
function customConfirm(message){
  return new Promise(resolve => {
    const overlay = el(`
      <div class="modal-overlay">
        <div class="modal-box">
          <div class="modal-message">${message}</div>
          <div class="modal-actions">
            <button class="btn btn-ghost modal-cancel">ยกเลิก</button>
            <button class="btn btn-primary modal-ok">ยืนยัน</button>
          </div>
        </div>
      </div>`);
    overlay.querySelector('.modal-cancel').onclick = () => { overlay.remove(); resolve(false); };
    overlay.querySelector('.modal-ok').onclick = () => { overlay.remove(); resolve(true); };
    document.body.appendChild(overlay);
  });
}

function customPrompt(message, placeholder){
  return new Promise(resolve => {
    const overlay = el(`
      <div class="modal-overlay">
        <div class="modal-box">
          <div class="modal-message">${message}</div>
          <input type="text" class="modal-input" placeholder="${placeholder || ''}">
          <div class="modal-actions">
            <button class="btn btn-ghost modal-cancel">ข้าม</button>
            <button class="btn btn-primary modal-ok">ตกลง</button>
          </div>
        </div>
      </div>`);
    const input = overlay.querySelector('.modal-input');
    overlay.querySelector('.modal-cancel').onclick = () => { overlay.remove(); resolve(null); };
    overlay.querySelector('.modal-ok').onclick = () => { overlay.remove(); resolve(input.value.trim()); };
    document.body.appendChild(overlay);
    setTimeout(() => input.focus(), 50);
  });
}




/* ============================================================
   HELPERS
   ============================================================ */
function gs(fn, ...args){
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler(resultStr => {
        try { resolve(resultStr == null ? null : JSON.parse(resultStr)); }
        catch(e){ reject(e); }
      })
      .withFailureHandler(err => { toast(err.message || 'เกิดข้อผิดพลาด'); reject(err); })
      .callServer(fn, JSON.stringify(args));
  });
}
function el(html){ const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstChild; }
function money(n){ return Number(n||0).toLocaleString('th-TH', {minimumFractionDigits:2, maximumFractionDigits:2}); }
function baht(n){
  n = Number(n||0);
  return (n < 0 ? '-฿' : '฿') + money(Math.abs(n));
}

function fmtDate(d){
  if(!d) return '-';
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString('th-TH', {day:'numeric', month:'short', year:'2-digit'});
}
let toastTimer;
function toast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=>t.classList.remove('show'), 2200);
}
function statusClass(emoji){
  if (emoji === '🟢') return 'green';
  if (emoji === '🟡') return 'amber';
  return 'red';
}

const content = document.getElementById('content');
const topbar  = document.getElementById('topbar');
const bottomNav = document.getElementById('bottomNav');

const NAV_TABS = [
  { key:'home',           icon:'🏠', label:'สินค้า',        hash:'#/home' },
  { key:'add-purchase',   icon:'🛒', label:'บันทึกซื้อ',    hash:'#/add-purchase' },
  { key:'pricing-picker', icon:'💰', label:'ตั้งราคาขาย',   hash:'#/pricing-picker' },
  { key:'price-history',  icon:'📊', label:'ประวัติราคา',   hash:'#/price-history' },
  { key:'master',         icon:'⚙️', label:'จัดการข้อมูล', hash:'#/master' }
];

function renderBottomNav(activeKey){
  bottomNav.innerHTML = '';
  NAV_TABS.forEach(t => {
    const btn = el(`<button class="nav-tab ${activeKey===t.key?'active':''}">
      <span class="nav-icon">${t.icon}</span><span>${t.label}</span>
    </button>`);
    btn.onclick = () => goTab(t.hash);
    bottomNav.appendChild(btn);
  });
}

function goTab(hash){
  navStack = [];
  go(hash);
}


let navStack = []; // { hash } for back button


/* ============================================================
   ROUTER
   ============================================================ */
window.addEventListener('hashchange', () => route(true));

window.addEventListener('DOMContentLoaded', () => {
  if (window.__initialPage === 'product' && window.__initialId){
    location.hash = '#/product/' + window.__initialId;
  }
  route(false);
  initPriceCompareFab();
});



function go(hash){ location.hash = hash; }
function goBack(){
  if (navStack.length > 1){
    navStack.pop();
    history.replaceState(null,'',navStack[navStack.length-1]);
    route(false);
  } else {
    go('#/home');
  }
}

function route(fromHashChange){
  const hash = location.hash || '#/home';
  if (!fromHashChange || navStack[navStack.length-1] !== hash){
    navStack.push(hash);
  }
  const parts = hash.replace('#/','').split('/');
  const [view, a, b] = parts;

  const navKeyMap = {
    home:'home', product:'home', store:'home', purchase:'home',
    'add-purchase':'add-purchase',
    pricing:'pricing-picker', 'pricing-picker':'pricing-picker',
    'price-history':'price-history',
    master:'master', 'store-ranking':'master', dashboard:'home', 'switch-store':'master'

  };
  renderBottomNav(navKeyMap[view] || 'home');

  let result;
  try{
    if (view === 'home' || !view)              result = renderHome();
    else if (view === 'product')                result = renderProduct(a);
    else if (view === 'store')                  result = renderStore(a, b);
    else if (view === 'purchase')                result = b === 'edit' ? renderEditPurchase(a) : renderPurchaseDetail(a);
    else if (view === 'pricing')                result = renderPricing(a);
    else if (view === 'pricing-picker')         result = renderPricingPicker();
    else if (view === 'add-purchase')           result = renderAddPurchase();
   else if (view === 'master')                 result = renderMaster();
   else if (view === 'store-ranking')          result = renderStoreRanking();
   else if (view === 'switch-store')           result = renderSwitchStore();   // ← เพิ่มบรรทัดนี้
   else if (view === 'reorder-list')           result = renderReorderList();

   else if (view === 'price-history')          result = a ? renderPriceHistoryDetail(a) : renderPriceHistoryPicker();
    else if (view === 'dashboard')          result = renderDashboard();
    else                                          result = renderHome();
  } catch(err){
    showFatalError(err);
    return;
  }
  Promise.resolve(result).catch(showFatalError);
}

function showFatalError(err){
  console.error(err);
  const msg = (err && err.message) ? err.message : String(err);
  content.innerHTML = `
    <div class="info-block" style="border-color:var(--danger); border-width:1.5px;">
      <div class="info-block-title" style="color:var(--danger);">⚠️ เกิดข้อผิดพลาด</div>
      <div style="font-size:13.5px; color:var(--ink-soft); white-space:pre-wrap; word-break:break-word; line-height:1.6;">${msg}</div>
    </div>
    <button class="btn btn-ghost" onclick="location.hash='#/home'; location.reload();">🏠 กลับหน้าแรก</button>
  `;
}

function setTopbar({ title, showBack, extra }){
  topbar.innerHTML = '';
  const row = el(`<div class="topbar-row">
    ${showBack ? `<button class="back-btn" id="backBtn">←</button>` : ''}
    <h1>${title}</h1>
  </div>`);
  topbar.appendChild(row);
  if (showBack) document.getElementById('backBtn').onclick = goBack;
  if (extra) topbar.appendChild(el(extra));
}





/* ============================================================
   HOME — 🏠 สินค้า
   ============================================================ */
let homeState = { categoryId: null, keyword: '' };

async function renderHome(){
  setTopbar({
    title: 'สินค้า', showBack: false,
    extra: `
     <div>
  <div class="searchbox">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#3A2E28" stroke-width="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    <input id="homeSearch" placeholder="ค้นหากลุ่มสินค้า..." value="${homeState.keyword}">
    <button id="scanQrBtn" style="border:none; background:transparent; font-size:19px; flex-shrink:0; padding:0 2px;">📷</button>
  </div>
  <div class="tabs" id="catTabs"></div>
</div>
`
});
  document.getElementById('scanQrBtn').onclick = () => openQRScannerExternal();
  document.getElementById('homeSearch').addEventListener('input', e => {
    homeState.keyword = e.target.value;
    renderProductGrid();
  });


  content.innerHTML = `<div id="homeSummary"></div><div id="productListWrap"><div class="spinner"></div></div>`;
  loadHomeSummary();

  // ดึงหมวดหมู่ + การ์ดสินค้า "พร้อมกัน" แทนรอทีละอย่าง
  const [categories] = await Promise.all([ gs('getCategories'), renderProductGrid() ]);
  const tabsEl = document.getElementById('catTabs');

   tabsEl.innerHTML = '';
  tabsEl.appendChild(el(`<button class="tab ${!homeState.categoryId ? 'active' : ''}" data-id="">ทั้งหมด</button>`));
  categories.forEach(c => {
    tabsEl.appendChild(el(`<button class="tab ${homeState.categoryId===c.id?'active':''}" data-id="${c.id}">${c.name}</button>`));
  });

  tabsEl.querySelectorAll('.tab').forEach(btn => {
    btn.onclick = () => {
      homeState.categoryId = btn.dataset.id || null;
      tabsEl.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
          renderProductGrid();
    };
  });
  window._categories = categories;
}


/* ============================================================
   ⚖️ ปุ่มลอยเทียบราคา — เทียบสินค้าที่ยังไม่ได้ซื้อ กับราคาที่มีอยู่ในระบบ
   กรอกแค่ ร้าน/ราคา/จำนวน ระบบดึงหน่วย+เรโชของสินค้าจากประวัติที่มีอยู่ให้อัตโนมัติ
   ============================================================ */
function initPriceCompareFab(){
  if (document.getElementById('priceCompareFab')) return; // กันสร้างซ้ำ

  const fab = el(`
    <button id="priceCompareFab" style="position:fixed; bottom:78px; right:16px; width:52px; height:52px; border-radius:50%; background:var(--orange); color:#fff; border:none; font-size:22px; box-shadow:0 4px 14px rgba(0,0,0,0.25); z-index:60; display:flex; align-items:center; justify-content:center;">⚖️</button>
  `);
  document.body.appendChild(fab);

  const panel = el(`
    <div id="priceComparePanel" style="position:fixed; left:0; right:0; bottom:0; background:var(--card); border-radius:20px 20px 0 0; box-shadow:0 -4px 20px rgba(0,0,0,0.2); z-index:61; padding:18px; padding-bottom:28px; transform:translateY(100%); transition:transform .25s ease; max-height:80vh; overflow-y:auto;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
        <div style="font-weight:700; font-size:16px;">⚖️ เทียบราคาสินค้า</div>
        <button id="pcCloseBtn" style="border:none; background:none; font-size:20px; color:var(--ink-soft);">✕</button>
      </div>

      <div class="field autocomplete">
        <label>สินค้าที่จะเทียบ</label>
        <input id="pcGroup" placeholder="พิมพ์ค้นหาสินค้า..." autocomplete="off">
        <div class="autocomplete-list" id="pcGroupList" style="display:none;"></div>
      </div>

      <div class="field"><label>ร้านค้าที่เจอราคานี้</label><input id="pcStore" placeholder="เช่น ชื่อร้าน">shop</div>
      <div class="field-row">
        <div class="field"><label>ราคาที่เจอ (บาท)</label><input type="number" id="pcPrice" inputmode="decimal"></div>
        <div class="field"><label id="pcQtyLabel">จำนวน</label><input type="number" id="pcQty" inputmode="decimal"></div>
      </div>

      <button class="btn btn-primary" id="pcCalcBtn" style="margin-top:6px;">เทียบราคา</button>
      <div id="pcResult" style="margin-top:14px;"></div>
    </div>
  `);
  document.body.appendChild(panel);

  let open = false;
  function togglePanel(show){
    open = show;
    panel.style.transform = show ? 'translateY(0)' : 'translateY(100%)';
  }
  fab.onclick = () => togglePanel(!open);
  panel.querySelector('#pcCloseBtn').onclick = () => togglePanel(false);

  setupAutocomplete('pcGroup', 'pcGroupList', 'productGroup', () => null);

  let refInfo = null; // {ratio, unit_buy, unit_sell, stores: [...]}

  document.getElementById('pcGroup').addEventListener('blur', async () => {
    const name = document.getElementById('pcGroup').value.trim();
    if (!name) return;
    const matches = await gs('searchMasterFor', 'productGroup', name, null);
    const exact = matches.find(m => m.name === name);
    if (!exact){ refInfo = null; return; }

    const stores = await gs('getStoreListByProductGroup', exact.id);
    if (stores.length === 0){ toast('สินค้านี้ยังไม่มีประวัติซื้อ เทียบราคาไม่ได้'); refInfo = null; return; }
    const latest = stores[0]; // getStoreListByProductGroup เรียงใหม่->เก่าอยู่แล้ว
    const ratio = latest.cost_per_buy_unit / latest.cost_per_sell_unit;
    refInfo = { ratio, unit_buy: latest.unit_buy, unit_sell: latest.unit_sell, stores };
    document.getElementById('pcQtyLabel').textContent = `จำนวน (${latest.unit_buy})`;
  });

  document.getElementById('pcCalcBtn').onclick = () => {
    const resultEl = document.getElementById('pcResult');
    if (!refInfo){ toast('กรุณาเลือกสินค้าที่มีในระบบก่อน'); return; }
    const price = parseFloat(document.getElementById('pcPrice').value || '0');
    const qty = parseFloat(document.getElementById('pcQty').value || '0');
    const storeName = document.getElementById('pcStore').value.trim() || 'ร้านนี้';
    if (!price || !qty){ toast('กรอกราคาและจำนวนให้ครบ'); return; }

    const candidateCostPerSellUnit = Math.round((price / qty / refInfo.ratio) * 100) / 100;
    const cheapest = refInfo.stores.reduce((min, s) => s.cost_per_sell_unit < min.cost_per_sell_unit ? s : min, refInfo.stores[0]);
    const latest = refInfo.stores[0];

    const diffVsCheapest = Math.round((candidateCostPerSellUnit - cheapest.cost_per_sell_unit) * 100) / 100;
    const pctVsCheapest = cheapest.cost_per_sell_unit ? Math.round((diffVsCheapest / cheapest.cost_per_sell_unit) * 1000) / 10 : 0;

    const verdict = diffVsCheapest < 0
      ? `🟢 ถูกกว่าที่เคยเจอถูกสุด ${Math.abs(diffVsCheapest)} บาท/${refInfo.unit_sell} (${Math.abs(pctVsCheapest)}%) — คุ้มค่าซื้อ!`
      : diffVsCheapest === 0
      ? `🟡 ราคาเท่ากับที่ถูกสุดที่เคยเจอพอดี`
      : `🔴 แพงกว่าที่เคยเจอถูกสุด ${diffVsCheapest} บาท/${refInfo.unit_sell} (${pctVsCheapest}%) — ที่ ${cheapest.store_name} เคยถูกกว่า`;

    resultEl.innerHTML = `
      <div class="info-block">
        <div class="info-row"><span class="k">ราคาที่ ${storeName}</span><span class="v num">${candidateCostPerSellUnit} บาท/${refInfo.unit_sell}</span></div>
        <div class="info-row"><span class="k">ถูกสุดที่เคยเจอ</span><span class="v num">${cheapest.cost_per_sell_unit} บาท/${refInfo.unit_sell} (${cheapest.store_name})</span></div>
        <div class="info-row"><span class="k">ล่าสุดที่ซื้อ</span><span class="v num">${latest.cost_per_sell_unit} บาท/${refInfo.unit_sell} (${latest.store_name})</span></div>
      </div>
      <div style="margin-top:8px; font-size:14px; line-height:1.6;">${verdict}</div>
    `;
  };
}



async function renderProductGrid(){
  const wrap = document.getElementById('productListWrap');
  wrap.innerHTML = `<div class="spinner"></div>`;

  let cards = await gs('getProductGroupCardsBatch', homeState.categoryId ? [homeState.categoryId] : null);

  if (homeState.keyword.trim()){
    const kw = homeState.keyword.trim().toLowerCase();
    cards = cards.filter(c => c.name.toLowerCase().includes(kw));
  }
  if (cards.length === 0){
    wrap.innerHTML = emptyState('🔍','ไม่พบกลุ่มสินค้า','ลองคำค้นอื่น หรือบันทึกการซื้อใหม่');
    return;
  }
  const list = el(`<div></div>`);
  cards.forEach(c => list.appendChild(productCard(c)));
  wrap.innerHTML = '';
  wrap.appendChild(list);
}



async function loadHomeSummary(){
  const wrap = document.getElementById('homeSummary');
  if (!wrap) return;
  try{
    const s = await gs('getMonthlySavings');
    wrap.innerHTML = '';

    // เดือนนี้ยังไม่มีรายการซื้อ — ยังคงต้องมีทางเข้า Dashboard เสมอ
    if (s.purchase_count === 0){
      const card = el(`
        <div class="info-block" style="margin-bottom:14px; cursor:pointer; text-align:center; color:var(--ink-soft);">
          <div style="font-size:13.5px;">📊 เดือนนี้ยังไม่มีรายการซื้อ — ดู Dashboard ›</div>
        </div>`);
      card.onclick = () => go('#/dashboard');
      wrap.appendChild(card);
      return;
    }

    const card = el(`
    <div>
      <div class="info-block" style="margin-bottom:14px; cursor:pointer;">
        <div class="info-block-title">💰 เดือนนี้ <span style="font-weight:400; color:var(--ink-soft); font-size:12px; margin-left:auto;">ดู Dashboard ›</span></div>
        <div class="info-row"><span class="k">ใช้จ่ายรวม</span><span class="v num">${baht(s.total_spend)}</span></div>
        <div class="info-row"><span class="k">ซื้อไป</span><span class="v num">${s.purchase_count} ครั้ง · ${s.store_count} ร้าน</span></div>
        ${s.total_saved > 0 ? `<div class="info-row total"><span class="k">🌱 ประหยัดได้</span><span class="v num" style="color:var(--mint-dk);">${baht(s.total_saved)}</span></div>` : ''}
      </div>
      </div>`);
    card.onclick = () => go('#/dashboard');
    wrap.appendChild(card);
  } catch(err){ /* ไม่ต้องแสดง error ถ้าโหลดสรุปไม่สำเร็จ ไม่กระทบการใช้งานหลัก */ }
}




// การ์ดแสดงกลุ่มสินค้าหน้าแรก  ลองตัด accept แล้วรันผ่าน ------
function productCard(c){
  const freqText = `ซื้อแล้ว ${c.purchase_count} ครั้ง`;

  const card = el(`
    <div class="product-card">
      <div class="tag-hole"></div>
      <div class="product-card-media">
        ${c.image_url ? `<img src="${c.image_url}">` : '<span class="product-card-emoji">🏷️</span>'}
        <input type="file" accept="image" style="display:none;" class="pc-file-input">
        <button class="icon-btn pc-upload-btn">📷</button>
      </div>
      <div class="product-card-info">
       <div class="product-card-name">${c.name} ${c.should_reorder ? '<span class="badge badge-amber">⏰ ควรซื้อเพิ่ม</span>' : ''}</div>

        <div class="product-card-date">🗓️ ${c.last_purchase_date ? fmtDate(c.last_purchase_date) : 'ยังไม่เคยซื้อ'}</div>
        <div class="product-card-cost-row">
          <div class="cost-chip cost-buy"><span class="cost-value num">${baht(c.cost_per_buy_unit)}</span><span class="cost-unit">/${c.unit_buy || '-'}</span></div>
          <div class="cost-chip cost-sell"><span class="cost-value num">${baht(c.cost_per_sell_unit)}</span><span class="cost-unit">/${c.unit_sell || '-'}</span></div>
        </div>
        <div class="product-card-sell-row">ขาย ${baht(c.sell_price_per_buy_unit)}/${c.unit_buy || '-'} · ${baht(c.sell_price_per_sell_unit)}/${c.unit_sell || '-'}</div>
        <div class="product-card-meta">🏪 ${c.store_count} ร้าน · 🔁 ${freqText}</div>
      </div>
    </div>`);

  card.onclick = () => go(`#/product/${c.id}`);

  const fileInput = card.querySelector('.pc-file-input');
  const uploadBtn = card.querySelector('.pc-upload-btn');
  uploadBtn.onclick = (e) => { e.stopPropagation(); fileInput.click(); };
  fileInput.onclick = (e) => e.stopPropagation();
  fileInput.onchange = (e) => {
    e.stopPropagation();
    const file = fileInput.files[0];
    handleProductImageUpload(c.id, file, () => renderProductGrid());
  };

  return card;
}


function emptyState(icon, title, sub){
  return `<div class="empty-state"><div class="icon">${icon}</div><div class="title">${title}</div><div>${sub}</div></div>`;
}


/* ============================================================
   PRODUCT DETAIL — รายละเอียดกลุ่มสินค้า + รายชื่อร้านค้า
   ============================================================ */
async function renderProduct(pgId){
  setTopbar({ title: 'รายละเอียดสินค้า', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const d = await gs('getProductGroupDetail', pgId);
  if (!d) throw new Error('ไม่พบข้อมูลกลุ่มสินค้านี้ (pgId: ' + pgId + ') — อาจถูกลบไปแล้ว หรือ id ไม่ถูกต้อง');
  const sc = statusClass(d.status.emoji);

content.innerHTML = '';
content.appendChild(el(`
  <div>
  <div class="tag-card">
    <div class="tag-hole"></div>
    <div class="tag-media" id="productMedia" style="position:relative;">
      ${d.image_url ? `<img src="${d.image_url}">` : '🏷️'}
      <input type="file" id="imgInput" accept="image/*" style="display:none;">
      <button class="icon-btn" id="imgBtn" style="position:absolute; bottom:8px; right:8px; background:rgba(255,255,255,0.9); box-shadow:var(--shadow);">📷</button>
      ${d.image_url ? `<button class="icon-btn danger" id="imgRemoveBtn" style="position:absolute; bottom:8px; right:50px; background:rgba(255,255,255,0.9); box-shadow:var(--shadow);">🗑️</button>` : ''}
    </div>
    <div class="tag-body"><div class="tag-name" style="font-size:19px;">${d.name}</div></div>
  </div>

  ${d.frequency && d.frequency.should_reorder ? `
  <div class="status-banner status-yellow">
    <span style="font-size:20px;">⏰</span>
    <div>ถึงเวลาสั่งซื้อเพิ่มแล้ว — ซื้อล่าสุดเมื่อ ${d.frequency.days_since_last_purchase} วันก่อน (ปกติซื้อทุก ~${d.frequency.avg_interval_days} วัน)</div>
  </div>` : ''}

  <div class="grid">
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ต้นทุน/หน่วยซื้อ</div><div class="stat-value num">${baht(d.cost_per_buy_unit)}<span style="font-size:11px; color:var(--ink-soft); font-weight:600;"> /${d.unit_buy || '-'}</span></div></div>
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ต้นทุน/หน่วยขาย</div><div class="stat-value num">${baht(d.cost_per_unit)}<span style="font-size:11px; color:var(--ink-soft); font-weight:600;"> /${d.unit_sell || '-'}</span></div></div>
  </div>

  <div class="grid" style="margin-top:10px; opacity:.8;">
    <div class="stat-tag" style="padding:10px 12px;"><div class="stat-label" style="font-size:11px;">ราคาขาย/หน่วยขาย</div><div class="stat-value num" style="font-size:15px;">${baht(d.current_sell_price)}</div></div>
    <div class="stat-tag" style="padding:10px 12px;"><div class="stat-label" style="font-size:11px;">กำไร/หน่วย · Margin</div><div class="stat-value num" style="font-size:15px;">${baht(d.profit_per_unit)} · ${d.margin_percent}%</div></div>
  </div>

  ${d.frequency && d.frequency.avg_interval_days ? `
  <div class="stat-tag" style="margin-top:10px; padding:10px 12px;">
    <div class="tag-hole"></div>
    <div class="stat-label" style="font-size:11px;">🔁 ความถี่การซื้อ</div>
    <div class="stat-value" style="font-size:14px;">ซื้อทุก ~${d.frequency.avg_interval_days} วัน · ล่าสุด ${d.frequency.days_since_last_purchase} วันก่อน</div>
  </div>` : ''}

  <div style="display:flex; gap:8px; margin-top:12px;">
    <button id="setPriceBtn" style="flex:1; padding:9px; border-radius:999px; border:1.5px solid var(--line); background:var(--card); font-size:12.5px; font-weight:700; color:var(--ink-soft);">💰 ตั้งราคาขาย</button>
    <button id="viewHistoryBtn" style="flex:1; padding:9px; border-radius:999px; border:1.5px solid var(--line); background:var(--card); font-size:12.5px; font-weight:700; color:var(--ink-soft);">📊 ประวัติราคา</button>
    <button id="qrBtn" style="flex-shrink:0; width:42px; padding:9px; border-radius:999px; border:1.5px solid var(--line); background:var(--card); font-size:15px;">🔲</button>
  </div>

  <div class="section-title">🏪 ร้านค้า (เรียงจากซื้อล่าสุด)</div>
  <div id="storeList"></div>
  </div>
`));


  const imgInput = document.getElementById('imgInput');
  document.getElementById('imgBtn').onclick = () => imgInput.click();
  imgInput.onchange = () => {
    const file = imgInput.files[0];
    handleProductImageUpload(pgId, file, (url) => {
      document.getElementById('productMedia').querySelector('img,span')?.remove();
      renderProduct(pgId); // โหลดใหม่เพื่อแสดงรูปล่าสุดและปุ่มลบรูป
    });
  };
  const removeBtn = document.getElementById('imgRemoveBtn');
  if (removeBtn){
    removeBtn.onclick = async () => {
      if (!(await customConfirm('ลบรูปนี้?'))) return;
      try{ await gs('removeProductGroupImage', pgId); toast('ลบรูปแล้ว'); }
      finally { renderProduct(pgId); }
    };
  }

document.getElementById('setPriceBtn').onclick = () => go(`#/pricing/${pgId}`);
document.getElementById('viewHistoryBtn').onclick = () => go(`#/price-history/${pgId}`);
document.getElementById('qrBtn').onclick = () => showProductQR(pgId, d.name);


  const list = document.getElementById('storeList');
  if (d.stores.length === 0){
    list.innerHTML = emptyState('🏪','ยังไม่มีประวัติร้านค้า','');
  }
  d.stores.forEach(s => {
    const badges = [];
    if (s.is_cheapest) badges.push('<span class="badge badge-mint">🌱 ถูกสุด</span>');
    if (s.is_most_expensive) badges.push('<span class="badge badge-red">🔺 แพงสุด</span>');
    if (s.is_latest) badges.push('<span class="badge badge-amber">🕐 ล่าสุด</span>');

    const row = el(`
      <div class="row-card">
        <div class="avatar">${s.store_name.charAt(0)}</div>
        <div class="row-main">
          <div class="row-title">${s.store_name} ${badges.join(' ')}</div>
          <div class="row-sub">🗓️ ${fmtDate(s.last_purchase_date)} · ซื้อแล้ว ${s.purchase_count} ครั้ง</div>
          <div class="row-sub num" style="margin-top:2px;">${baht(s.cost_per_buy_unit)}/${s.unit_buy} · ${baht(s.cost_per_sell_unit)}/${s.unit_sell}</div>
        </div>
        <div class="row-arrow">›</div>
      </div>`);
    row.onclick = () => go(`#/store/${pgId}/${s.store_id}`);
    list.appendChild(row);
  });
}


/* ============================================================
   STORE DETAIL — รายละเอียดร้านค้า + ประวัติการซื้อ
   ============================================================ */
async function renderStore(pgId, storeId){
  setTopbar({ title: 'รายละเอียดร้านค้า', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;
  await renderStoreList(pgId, storeId, false);
}


async function renderStoreList(pgId, storeId, includeCancelled){
  const d = await gs('getStoreDetail', pgId, storeId, includeCancelled);
  const s = d.summary;

 content.innerHTML = '';
 content.appendChild(el(`
  <div>
  <div class="tag-card"><div class="tag-hole"></div>
    <div class="tag-body" style="padding-top:16px;"><div class="tag-name" style="font-size:19px;">🏪 ${d.store_name}</div></div>
  </div>

  <div class="grid">
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ราคาล่าสุด</div><div class="stat-value num">${baht(s?.latest_price)}</div></div>
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ซื้อล่าสุด</div><div class="stat-value">${fmtDate(s?.latest_date)}</div></div>
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ซื้อทั้งหมด</div><div class="stat-value num">${s?.purchase_count||0} ครั้ง</div></div>
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">เฉลี่ย</div><div class="stat-value num">${baht(s?.avg_price)}</div></div>
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ต่ำสุด</div><div class="stat-value num" style="color:var(--mint-dk)">${baht(s?.min_price)}</div></div>
    <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">สูงสุด</div><div class="stat-value num" style="color:var(--danger)">${baht(s?.max_price)}</div></div>
  </div>

  <div class="section-title" style="display:flex; justify-content:space-between; align-items:center;">
    <span>🧾 ประวัติการซื้อ (ใหม่ → เก่า)</span>
    <label style="display:flex; align-items:center; gap:6px; text-transform:none; letter-spacing:0; font-size:12.5px; font-weight:600;">
      <input type="checkbox" id="showCancelledCk" ${includeCancelled ? 'checked' : ''}> แสดงรายการที่ยกเลิก
    </label>
  </div>
  <div id="purchaseList"></div>
  </div>
`));


  document.getElementById('showCancelledCk').onchange = (e) => renderStoreList(pgId, storeId, e.target.checked);

   const list = document.getElementById('purchaseList');
  d.purchases.forEach(p => {
    const isCancelled = p.status === 'cancelled';
    const row = el(`
    <div>
      <div class="row-card ${isCancelled ? 'cancelled' : ''}">
        <div class="avatar" style="background:${isCancelled ? 'var(--line-dash)' : 'var(--mint)'};">${isCancelled ? '🚫' : '🧾'}</div>
        <div class="row-main">
          <div class="row-title">${fmtDate(p.date)} ${isCancelled ? '<span class="badge" style="background:var(--line); color:var(--ink-soft);">ยกเลิก</span>' : ''}</div>
          <div class="row-sub">${p.qty_buy} ${p.unit_buy} · ${baht(p.price_buy)}</div>
        </div>
        <div class="row-price num">${baht(p.cost_per_unit)}</div>
        <div class="row-arrow">›</div>
      </div>
      </div>`);
    row.onclick = () => go(`#/purchase/${p.id}`);
    list.appendChild(row);
  });
  if (d.purchases.length === 0) list.innerHTML = emptyState('🧾','ยังไม่มีประวัติการซื้อ','');
}


/* ============================================================
   PURCHASE DETAIL — 4 ส่วน
   ============================================================ */
async function renderPurchaseDetail(purchaseId){
  setTopbar({ title: 'รายละเอียดการซื้อ', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const p = await gs('getPurchaseDetail', purchaseId);
  const isCancelled = p.status === 'cancelled';

  content.innerHTML = `
    ${isCancelled ? `
    <div class="status-banner status-red">
      <span style="font-size:20px;">🚫</span>
      <div>รายการนี้ถูกยกเลิกแล้ว${p.cancelled_at ? ' เมื่อ ' + fmtDate(p.cancelled_at) : ''}${p.cancel_reason ? ' — ' + p.cancel_reason : ''}</div>
    </div>` : ''}

    <div class="info-block">
      <div class="info-block-title">📋 ข้อมูลการซื้อ</div>
      <div class="info-row"><span class="k">วันที่ซื้อ</span><span class="v">${fmtDate(p.date)}</span></div>
      <div class="info-row"><span class="k">ร้านค้า</span><span class="v">${p.store_name}</span></div>
      <div class="info-row"><span class="k">เลขที่คำสั่งซื้อ</span><span class="v">${p.order_no || '-'}</span></div>
<div class="info-row"><span class="k">ลิงค์สินค้า</span><span class="v">${p.product_url ? `<a href="${p.product_url}" target="_blank" rel="noopener" style="color:var(--orange-dk);">เปิดลิงก์</a>` : '-'}</span></div>
<div class="info-row"><span class="k">ผู้บันทึก</span><span class="v">${p.recorder || '-'}</span></div>

      <div class="info-row"><span class="k">หมายเหตุ</span><span class="v">${p.note || '-'}</span></div>
    </div>

    <div class="info-block">
      <div class="info-block-title">📦 ข้อมูลสินค้า</div>
      <div class="info-row"><span class="k">ชื่อสินค้า</span><span class="v">${p.product_group_name}</span></div>
      <div class="info-row"><span class="k">จำนวนซื้อ</span><span class="v num">${p.qty_buy} ${p.unit_buy}</span></div>
      <div class="info-row"><span class="k">เรโช</span><span class="v num">1 : ${p.ratio}</span></div>
      <div class="info-row"><span class="k">จำนวนขายรวม</span><span class="v num">${p.qty_sell_total} ${p.unit_sell}</span></div>
    </div>

    <div class="info-block">
      <div class="info-block-title">💰 ข้อมูลต้นทุน</div>
      <div class="info-row"><span class="k">ราคาซื้อ</span><span class="v num">${baht(p.price_buy)}</span></div>
      <div class="info-row"><span class="k">ค่าขนส่ง</span><span class="v num">${baht(p.shipping)}</span></div>
      <div class="info-row"><span class="k">ส่วนลด</span><span class="v num">-${baht(p.discount)}</span></div>
      <div class="info-row"><span class="k">ค่าใช้จ่ายอื่นๆ</span><span class="v num">${baht(p.other_cost)}</span></div>
      <div class="info-row total"><span class="k">ต้นทุนสุทธิ</span><span class="v num">${baht(p.net_cost)}</span></div>
      <div class="info-row total"><span class="k">ราคาทุน/หน่วยขาย</span><span class="v num">${baht(p.cost_per_unit)}</span></div>
    </div>


      <div class="info-block">
      <div class="info-block-title">📎 เอกสาร</div>
      <div style="font-size:12.5px; color:var(--ink-soft); margin-bottom:6px;">รูปสินค้า</div>
      ${p.image_url ? `<a href="${p.image_url}" target="_blank" rel="noopener"><img class="receipt-thumb" src="${p.image_url}" style="margin-bottom:12px; cursor:pointer;"></a>` : `<div style="font-size:13px; color:var(--ink-soft); margin-bottom:12px;">— ยังไม่มีรูป (อัปโหลดได้ที่หน้ารายละเอียดกลุ่มสินค้า) —</div>`}

      <div style="font-size:12.5px; color:var(--ink-soft); margin-bottom:6px;">🧾 รูปใบเสร็จ</div>
            ${p.receipt_img ? `<a href="${p.receipt_img}" target="_blank" rel="noopener"><img class="receipt-thumb" src="${p.receipt_img}" style="cursor:pointer;"></a>` : `<div style="font-size:13px; color:var(--ink-soft);">— ยังไม่มีรูป —</div>`}


     <input type="file"  id="receiptInput" style="display:none;">
      <button type="button" class="btn btn-ghost" id="receiptBtn" style="margin-top:8px; margin-bottom:16px;">${p.receipt_img ? '📷 เปลี่ยนรูปใบเสร็จ' : '📷 อัปโหลดรูปใบเสร็จ (ไม่บังคับ)'}</button>

      <div style="font-size:12.5px; color:var(--ink-soft); margin-bottom:6px;">🧾 รูปบิล</div>
      ${p.bill_img ? `<a href="${p.bill_img}" target="_blank" rel="noopener"><img class="receipt-thumb" src="${p.bill_img}" style="cursor:pointer;"></a>` : `<div style="font-size:13px; color:var(--ink-soft);">— ยังไม่มีรูป —</div>`}
      <input type="file"  id="billInput" style="display:none;">
      <button type="button" class="btn btn-ghost" id="billBtn" style="margin-top:8px;">${p.bill_img ? '📷 เปลี่ยนรูปบิล' : '📷 อัปโหลดรูปบิล (ไม่บังคับ)'}</button>
    </div>


    ${!isCancelled ? `
      <button class="btn btn-ghost" id="editPurchaseBtn">✏️ แก้ไขรายการ</button>
      <button class="btn btn-ghost" id="cancelPurchaseBtn" style="margin-top:10px; color:var(--danger); border-color:var(--danger);">🚫 ยกเลิกรายการ</button>
    ` : `
      <button class="btn btn-mint" id="restorePurchaseBtn">↩️ กู้คืนรายการ</button>
    `}
  `;
  document.getElementById('receiptBtn').onclick = () => document.getElementById('receiptInput').click();
  document.getElementById('receiptInput').onchange = (e) => {
    handlePurchaseImageUpload(purchaseId, 'receipt_img', e.target.files[0], () => renderPurchaseDetail(purchaseId));
  };
  document.getElementById('billBtn').onclick = () => document.getElementById('billInput').click();
  document.getElementById('billInput').onchange = (e) => {
    handlePurchaseImageUpload(purchaseId, 'bill_img', e.target.files[0], () => renderPurchaseDetail(purchaseId));
  };

  const editBtn = document.getElementById('editPurchaseBtn');
  if (editBtn) editBtn.onclick = () => go(`#/purchase/${purchaseId}/edit`);

  const cancelBtn = document.getElementById('cancelPurchaseBtn');
  if (cancelBtn) cancelBtn.onclick = async () => {
    if (!(await customConfirm('ยกเลิกรายการซื้อนี้? (จะไม่ถูกนำมาคิดต้นทุน/ราคาอีก แต่ยังดูประวัติย้อนหลังได้)'))) return;
    const reasonInput = await customPrompt('เหตุผลที่ยกเลิก (ไม่บังคับ):', 'เช่น ซื้อผิด, ยกเลิกออเดอร์');
    const reason = reasonInput || '';
    cancelBtn.textContent = 'กำลังยกเลิก...'; cancelBtn.disabled = true;
    try{
      await gs('cancelPurchase', purchaseId, reason);
      toast('ยกเลิกรายการแล้ว');
      renderPurchaseDetail(purchaseId);
    } finally { cancelBtn.disabled = false; }
  };

  const restoreBtn = document.getElementById('restorePurchaseBtn');
  if (restoreBtn) restoreBtn.onclick = async () => {
    if (!(await customConfirm('กู้คืนรายการนี้กลับมาใช้งาน?'))) return;
    restoreBtn.textContent = 'กำลังกู้คืน...'; restoreBtn.disabled = true;
    try{
      await gs('restorePurchase', purchaseId);
      toast('กู้คืนรายการแล้ว');
      renderPurchaseDetail(purchaseId);
    } finally { restoreBtn.disabled = false; }
  };
}

/* ============================================================
   แก้ไขรายการซื้อ
   ============================================================ */
async function renderEditPurchase(purchaseId){
  setTopbar({ title: 'แก้ไขรายการซื้อ', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const p = await gs('getPurchaseDetail', purchaseId);

  content.innerHTML = `
    <div class="info-block" style="font-size:13.5px; color:var(--ink-soft);">
      📦 ${p.product_group_name}
    </div>

    <div class="field"><label>ร้านค้า</label><input id="fStore" value="${p.store_name}"></div>


<div class="field"><label>ลิงค์สินค้า (ไม่บังคับ)</label><input id="fProductUrl" type="url" value="${p.product_url || ''}" placeholder="เช่น ลิงก์ Shopee, Lazada"></div>





    <div class="field-row">
      <div class="field"><label>วันที่ซื้อ</label><input type="date" id="fDate" value="${p.date ? new Date(p.date).toISOString().slice(0,10) : ''}"></div>
      <div class="field"><label>เลขที่คำสั่งซื้อ</label><input id="fOrderNo" value="${p.order_no || ''}"></div>
    </div>

    <div class="section-title">📦 จำนวน & หน่วย</div>
    <div class="field-row">
      <div class="field"><label>จำนวนซื้อ</label><input type="number" id="fQty" inputmode="decimal" value="${p.qty_buy}"></div>
      <div class="field"><label>หน่วยซื้อ</label><input id="fUnitBuy" value="${p.unit_buy}"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>เรโช</label><input type="number" id="fRatio" inputmode="decimal" value="${p.ratio}"></div>
      <div class="field"><label>หน่วยขาย</label><input id="fUnitSell" value="${p.unit_sell}"></div>
    </div>

    <div class="section-title">💰 ต้นทุน</div>
    <div class="field-row">
      <div class="field"><label>ราคาซื้อ (บาท)</label><input type="number" id="fPrice" inputmode="decimal" value="${p.price_buy}"></div>
      <div class="field"><label>ค่าขนส่ง</label><input type="number" id="fShipping" inputmode="decimal" value="${p.shipping || 0}"></div>
    </div>
    <div class="field"><label>ค่าใช้จ่ายอื่นๆ</label><input type="number" id="fOther" inputmode="decimal" value="${p.other_cost || 0}"></div>

    <label style="display:block; font-size:13px; font-weight:600; color:var(--ink-soft); margin-bottom:6px;">ส่วนลด (ถ้ามี — เพิ่มได้หลายรายการ)</label>
    <div id="discountList"></div>
    <button type="button" class="btn btn-ghost" id="addDiscountBtn" style="margin-bottom:14px;">+ เพิ่มรายการส่วนลด</button>

    <div class="field"><label>หมายเหตุ</label><textarea id="fNote" rows="2">${p.note || ''}</textarea></div>

        <label style="display:block; font-size:13px; font-weight:600; color:var(--ink-soft); margin-bottom:6px;">📎 รูปเอกสาร (ไม่บังคับ)</label>
    <div style="font-size:12.5px; color:var(--ink-soft); margin-bottom:4px;">รูปใบเสร็จ</div>
    ${p.receipt_img ? `<a href="${p.receipt_img}" target="_blank" rel="noopener"><img class="receipt-thumb" src="${p.receipt_img}" style="cursor:pointer;"></a>` : ''}
    <input type="file" accept="image/*" id="receiptInput" style="display:none;">
    <button type="button" class="btn btn-ghost" id="receiptBtn" style="margin-top:6px; margin-bottom:14px;">${p.receipt_img ? '📷 เปลี่ยนรูปใบเสร็จ' : '📷 อัปโหลดรูปใบเสร็จ'}</button>

    <div style="font-size:12.5px; color:var(--ink-soft); margin-bottom:4px;">รูปบิล</div>
    ${p.bill_img ? `<a href="${p.bill_img}" target="_blank" rel="noopener"><img class="receipt-thumb" src="${p.bill_img}" style="cursor:pointer;"></a>` : ''}
    <input type="file" accept="image/*" id="billInput" style="display:none;">
    <button type="button" class="btn btn-ghost" id="billBtn" style="margin-top:6px; margin-bottom:16px;">${p.bill_img ? '📷 เปลี่ยนรูปบิล' : '📷 อัปโหลดรูปบิล'}</button>

    <button class="btn btn-primary" id="saveEditBtn">✔ บันทึกการแก้ไข</button>

  `;

  setupDiscountRows('discountList', 'addDiscountBtn', p.discounts);

  document.getElementById('receiptBtn').onclick = () => document.getElementById('receiptInput').click();
  document.getElementById('receiptInput').onchange = (e) => {
    handlePurchaseImageUpload(purchaseId, 'receipt_img', e.target.files[0], () => renderEditPurchase(purchaseId));
  };
  document.getElementById('billBtn').onclick = () => document.getElementById('billInput').click();
  document.getElementById('billInput').onchange = (e) => {
    handlePurchaseImageUpload(purchaseId, 'bill_img', e.target.files[0], () => renderEditPurchase(purchaseId));
  };

  document.getElementById('saveEditBtn').onclick = async () => {
   const payload = {
  store_name: document.getElementById('fStore').value.trim(),
  product_url: document.getElementById('fProductUrl')?.value.trim() || '',
  date: document.getElementById('fDate').value,
  order_no: document.getElementById('fOrderNo').value.trim(),
  qty_buy: parseFloat(document.getElementById('fQty').value || '0'),
  unit_buy: document.getElementById('fUnitBuy').value.trim(),
  ratio: parseFloat(document.getElementById('fRatio').value || '0'),
  unit_sell: document.getElementById('fUnitSell').value.trim(),
  price_buy: parseFloat(document.getElementById('fPrice').value || '0'),
  shipping: parseFloat(document.getElementById('fShipping').value || '0'),
  other_cost: parseFloat(document.getElementById('fOther').value || '0'),
  discounts: collectDiscounts('discountList', document.getElementById('fPrice').value),
  note: document.getElementById('fNote').value.trim()
};

    const btn = document.getElementById('saveEditBtn');
    btn.textContent = 'กำลังบันทึก...'; btn.disabled = true;
    try{
      await gs('updatePurchase', purchaseId, payload);
      toast('บันทึกการแก้ไขแล้ว');
      go(`#/purchase/${purchaseId}`);
    } finally { btn.textContent = '✔ บันทึกการแก้ไข'; btn.disabled = false; }
  };
}


/* ============================================================
   ตั้งราคาขาย
   ============================================================ */
async function renderPricing(pgId){
  setTopbar({ title: 'ตั้งราคาขาย', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const info = await gs('getPricingInfo', pgId);
  const currentEmail = await gs('getCurrentUserEmail').catch(() => '');
  let mode = 'price'; // 'price' | 'margin'
  let roundInt = false;

  content.innerHTML = `
    <div class="info-block">
      <div class="info-row"><span class="k">ต้นทุน/หน่วยขาย</span><span class="v num">${baht(info.cost_per_unit)}</span></div>
      <div class="info-row"><span class="k">ราคาขายปัจจุบัน</span><span class="v num">${baht(info.current_sell_price)}</span></div>
    </div>

    <div class="segmented">
      <button id="modePrice" class="active">กรอกราคาขายเอง</button>
      <button id="modeMargin">ตั้งจาก % กำไร</button>
    </div>

    <div class="field" id="inputWrap">
      <label id="inputLabel">ราคาขาย (บาท)</label>
      <input type="number" id="mainInput" inputmode="decimal" step="0.01" value="${info.current_sell_price}">
    </div>

    <div class="checkbox-row">
      <input type="checkbox" id="roundInt"> <label for="roundInt" style="margin:0;">ปัดขึ้นเป็นจำนวนเต็ม (11.26 → 12)</label>

    </div>

    <div class="grid" style="margin-bottom:18px;">
      <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ราคาขาย</div><div class="stat-value num" id="outPrice">-</div></div>
      <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">กำไร/หน่วย</div><div class="stat-value num" id="outProfit">-</div></div>
      <div class="stat-tag" style="grid-column:1/-1;"><div class="tag-hole"></div><div class="stat-label">Margin</div><div class="stat-value num" id="outMargin">-</div></div>
    </div>

    <button class="btn btn-primary" id="savePriceBtn">✔ บันทึกราคาขาย</button>
  `;

  const mainInput = document.getElementById('mainInput');
  const inputLabel = document.getElementById('inputLabel');
  const roundCk = document.getElementById('roundInt');

  function recalc(){
    const val = parseFloat(mainInput.value || '0');
    let result;
    if (mode === 'price'){
      result = calcFromSellPriceLocal(info.cost_per_unit, val);
    } else {
      result = calcFromMarginLocal(info.cost_per_unit, val, roundCk.checked);
    }
    document.getElementById('outPrice').textContent = baht(result.sell_price);
    document.getElementById('outProfit').textContent = baht(result.profit_per_unit);
    document.getElementById('outMargin').textContent = result.margin_percent + '%';
    return result;
  }
  // ใช้ค่าจาก server formula ผ่าน RPC เพื่อความแม่นยำตรงกับฝั่ง Apps Script
  function calcFromSellPriceLocal(cost, sell){
    const profit = sell - cost;
    const margin = sell > 0 ? (profit/sell)*100 : 0;
    return { sell_price: sell, profit_per_unit: Math.round(profit*100)/100, margin_percent: Math.round(margin*10)/10 };
  }



    function calcFromMarginLocal(cost, marginPct, round){
    let sell = cost / (1 - marginPct/100);
    sell = round ? Math.ceil(sell) : Math.round(sell*100)/100;
    const profit = sell - cost;
    return { sell_price: sell, profit_per_unit: Math.round(profit*100)/100, margin_percent: marginPct };
  }

  mainInput.addEventListener('input', recalc);
  roundCk.addEventListener('change', recalc);

  document.getElementById('modePrice').onclick = () => {
    mode = 'price';
    document.getElementById('modePrice').classList.add('active');
    document.getElementById('modeMargin').classList.remove('active');
    inputLabel.textContent = 'ราคาขาย (บาท)';
    mainInput.value = info.current_sell_price;
    recalc();
  };
  document.getElementById('modeMargin').onclick = () => {
    mode = 'margin';
    document.getElementById('modeMargin').classList.add('active');
    document.getElementById('modePrice').classList.remove('active');
    inputLabel.textContent = 'กำไรที่ต้องการ (%)';
    mainInput.value = info.margin_percent;
    recalc();
  };

  recalc();

  document.getElementById('savePriceBtn').onclick = async () => {
    const result = recalc();
    if (!result.sell_price || result.sell_price <= 0){ toast('กรุณากรอกราคาที่ถูกต้อง'); return; }
    const btn = document.getElementById('savePriceBtn');
    btn.textContent = 'กำลังบันทึก...'; btn.disabled = true;
    try{
      await gs('saveSellPrice', { product_group_id: pgId, new_price: result.sell_price, edited_by: currentEmail });
      toast('บันทึกราคาขายแล้ว');
      go(`#/product/${pgId}`);
    } finally {
      btn.textContent = '✔ บันทึกราคาขาย'; btn.disabled = false;
    }
  };
}



/* ============================================================
   บันทึกการซื้อ — ฟอร์ม + autocomplete
   ============================================================ */
/* ============================================================
   บันทึกการซื้อ (รองรับหลายสินค้าใน 1 ออเดอร์)
   หารค่าส่ง/ค่าอื่นๆ/ส่วนลดรวม ตามสัดส่วนราคาสินค้าแต่ละรายการอัตโนมัติ
   ============================================================ */
let orderItemSeq = 0;

function createOrderItemRow(){
  const idx = orderItemSeq++;
  const row = el(`
    <div class="order-item" data-idx="${idx}" style="border:1.5px solid var(--line); border-radius:var(--radius-md); padding:14px; margin-bottom:12px; position:relative;">
      <button type="button" class="item-remove-btn" data-idx="${idx}" style="position:absolute; top:10px; right:10px; background:none; border:none; color:var(--danger); font-size:18px; cursor:pointer;">✕</button>
      <div style="font-size:12.5px; font-weight:700; color:var(--ink-soft); margin-bottom:8px;">สินค้าที่ ${idx+1}</div>

      <div class="field autocomplete">
        <label>หมวดสินค้าหลัก</label>
        <input id="fCategory-${idx}" placeholder="เช่น ขนม" autocomplete="off">
        <div class="autocomplete-list" id="listCategory-${idx}" style="display:none;"></div>
      </div>
      <div class="field autocomplete">
        <label>กลุ่มสินค้า</label>
        <input id="fGroup-${idx}" placeholder="เช่น มันฝรั่งทอด" autocomplete="off">
        <div class="autocomplete-list" id="listGroup-${idx}" style="display:none;"></div>
      </div>
      <div class="field"><label>ลิงค์สินค้า (ไม่บังคับ)</label><input id="fProductUrl-${idx}" type="url" placeholder="เช่น ลิงก์ Shopee, Lazada"></div>

      <div class="field-row">
        <div class="field"><label>จำนวนซื้อ</label><input type="number" id="fQty-${idx}" inputmode="decimal"></div>
        <div class="field"><label>หน่วยซื้อ</label><input id="fUnitBuy-${idx}" placeholder="ลัง"></div>
      </div>
      <div class="field-row">
        <div class="field"><label>เรโช (1 หน่วยซื้อ = กี่หน่วยขาย)</label><input type="number" id="fRatio-${idx}" inputmode="decimal"></div>
        <div class="field"><label>หน่วยขาย</label><input id="fUnitSell-${idx}" placeholder="ชิ้น"></div>
      </div>
            
      <div class="field"><label>ราคาต่อชิ้น (ไม่บังคับ — กรอกแล้วคำนวณราคารวมให้อัตโนมัติ)</label><input type="number" id="fUnitPrice-${idx}" inputmode="decimal"></div>
      <div class="field"><label>ราคาสินค้ารายการนี้ (บาท รวมทั้งจำนวนที่ซื้อ)</label><input type="number" id="fItemPrice-${idx}" inputmode="decimal"></div>

   
    </div>
  `);
  return { row, idx };
}

async function renderAddPurchase(){
  setTopbar({ title: 'บันทึกการซื้อ', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const today = new Date().toISOString().slice(0,10);
  const currentEmail = await gs('getCurrentUserEmail').catch(() => '');
  orderItemSeq = 0;



  content.innerHTML = `
    <div class="field autocomplete">
      <label>ร้านค้า</label>
      <input id="fStore" placeholder="เช่น Makro" autocomplete="off">
      <div class="autocomplete-list" id="listStore" style="display:none;"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>วันที่ซื้อ</label><input type="date" id="fDate" value="${today}"></div>
      <div class="field"><label>เลขที่คำสั่งซื้อ</label><input id="fOrderNo" placeholder="PO-1001"></div>
    </div>

    <button type="button" class="btn btn-ghost" id="ocrUploadBtn" style="margin-bottom:12px;">
  📸 อัพโหลดรูปคำสั่งซื้อ ให้ AI อ่านให้อัตโนมัติ
</button>

<div class="section-title">🛒 สินค้าในออเดอร์นี้</div>
<div id="orderItemsList"></div>

    <button type="button" class="btn btn-ghost" id="addItemBtn" style="margin-bottom:16px;">+ เพิ่มสินค้าอีกรายการ</button>

    <div class="section-title">💰 ค่าใช้จ่ายรวมทั้งออเดอร์ (หารตามสัดส่วนราคาสินค้าให้อัตโนมัติ)</div>
    <div class="field-row">
      <div class="field"><label>ค่าขนส่งรวม</label><input type="number" id="fShipping" inputmode="decimal" value="0"></div>
      <div class="field"><label>ค่าใช้จ่ายอื่นๆ รวม</label><input type="number" id="fOther" inputmode="decimal" value="0"></div>
    </div>

    <label style="display:block; font-size:13px; font-weight:600; color:var(--ink-soft); margin-bottom:6px;">ส่วนลดรวมทั้งออเดอร์ (ถ้ามี — เพิ่มได้หลายรายการ)</label>
    <div id="discountList"></div>
    <button type="button" class="btn btn-ghost" id="addDiscountBtn" style="margin-bottom:14px;">+ เพิ่มรายการส่วนลด</button>

    <div class="field"><label>ผู้บันทึก</label>
      ${currentEmail
        ? `<div style="padding:13px 14px; background:var(--bg); border:1.5px solid var(--line); border-radius:var(--radius-md); font-size:14.5px; color:var(--ink-soft); display:flex; align-items:center; gap:8px;">🧑 ${currentEmail}</div>`
        : `<input id="fRecorder" placeholder="กรอกชื่อผู้บันทึก">`}
    </div>
    <div class="field"><label>หมายเหตุ</label><textarea id="fNote" rows="2"></textarea></div>

    <div class="field-row">
      <div class="field">
        <label>รูปใบเสร็จ</label>
        <input type="file" accept="image/*" id="receiptInput" style="display:none;">
        <button type="button" class="btn btn-ghost" id="receiptBtn" style="width:100%;">📷 เลือกรูปใบเสร็จ</button>
        <div id="receiptFileName" style="font-size:11.5px; color:var(--ink-soft); margin-top:5px;"></div>
      </div>
      <div class="field">
        <label>รูปบิล</label>
        <input type="file" accept="image/*" id="billInput" style="display:none;">
        <button type="button" class="btn btn-ghost" id="billBtn" style="width:100%;">📷 เลือกรูปบิล</button>
        <div id="billFileName" style="font-size:11.5px; color:var(--ink-soft); margin-top:5px;"></div>
      </div>
    </div>

    <div id="allocPreview" style="font-size:12.5px; color:var(--ink-soft); margin:10px 0; line-height:1.6; background:var(--bg); border-radius:var(--radius-md); padding:10px 12px;"></div>

    <button class="btn btn-primary" id="submitBtn" style="margin-top:8px;">✔ บันทึกการซื้อทั้งออเดอร์</button>
  `;

  setupDiscountRows('discountList', 'addDiscountBtn', []);

  document.getElementById('receiptBtn').onclick = () => document.getElementById('receiptInput').click();
  document.getElementById('receiptInput').onchange = (e) => {
    document.getElementById('receiptFileName').textContent = e.target.files[0] ? '✓ ' + e.target.files[0].name : '';
  };
  document.getElementById('billBtn').onclick = () => document.getElementById('billInput').click();
  document.getElementById('billInput').onchange = (e) => {
    document.getElementById('billFileName').textContent = e.target.files[0] ? '✓ ' + e.target.files[0].name : '';
  };

  setupAutocomplete('fStore','listStore','store', () => null);

  const itemsList = document.getElementById('orderItemsList');

  






  function wireItemRow(idx){
    setupAutocomplete(`fCategory-${idx}`, `listCategory-${idx}`, 'category', () => null);
    setupAutocomplete(`fGroup-${idx}`, `listGroup-${idx}`, 'productGroup', () => null);
    

    // กรอกราคาต่อชิ้น + จำนวนซื้อ -> คำนวณราคารวมให้อัตโนมัติ (แก้ราคารวมเองทีหลังได้ ถ้าต้องการ override)
    function autoCalcTotal(){
      const unitPrice = parseFloat(document.getElementById(`fUnitPrice-${idx}`).value || '');
      const qty = parseFloat(document.getElementById(`fQty-${idx}`).value || '');
      if (!isNaN(unitPrice) && !isNaN(qty) && qty > 0){
        document.getElementById(`fItemPrice-${idx}`).value = Math.round(unitPrice * qty * 100) / 100;
        updateAllocPreview();
      }
    }
    document.getElementById(`fUnitPrice-${idx}`).addEventListener('input', autoCalcTotal);
    document.getElementById(`fQty-${idx}`).addEventListener('input', autoCalcTotal);

    document.getElementById(`fItemPrice-${idx}`).addEventListener('input', updateAllocPreview);
    itemsList.querySelector(`.item-remove-btn[data-idx="${idx}"]`).onclick = () => {
      if (itemsList.querySelectorAll('.order-item').length <= 1){ toast('ต้องมีสินค้าอย่างน้อย 1 รายการ'); return; }
      itemsList.querySelector(`.order-item[data-idx="${idx}"]`).remove();
      updateAllocPreview();
    };
  }
 
  window.wireItemRowGlobal = wireItemRow;   // ← เพิ่มบรรทัดนี้


  function addItemRow(){
    const { row, idx } = createOrderItemRow();
    itemsList.appendChild(row);
    wireItemRow(idx);
  }

  document.getElementById('addItemBtn').onclick = addItemRow;
  addItemRow(); // เริ่มต้นด้วย 1 รายการเสมอ
  document.getElementById('ocrUploadBtn').onclick = openOCRReaderExternal;
  document.getElementById('fShipping').addEventListener('input', updateAllocPreview);
  document.getElementById('fOther').addEventListener('input', updateAllocPreview);
  document.getElementById('discountList').addEventListener('input', updateAllocPreview);
  document.getElementById('addDiscountBtn').addEventListener('click', () => setTimeout(updateAllocPreview, 0));

  function getItemRows(){
    return Array.from(itemsList.querySelectorAll('.order-item')).map(r => {
      const idx = r.dataset.idx;
      return {
        idx,
        category_name: document.getElementById(`fCategory-${idx}`).value.trim(),
        product_group_name: document.getElementById(`fGroup-${idx}`).value.trim(),
        product_url: document.getElementById(`fProductUrl-${idx}`).value.trim(),
        qty_buy: parseFloat(document.getElementById(`fQty-${idx}`).value || '0'),
        unit_buy: document.getElementById(`fUnitBuy-${idx}`).value.trim(),
        ratio: parseFloat(document.getElementById(`fRatio-${idx}`).value || '0'),
        unit_sell: document.getElementById(`fUnitSell-${idx}`).value.trim(),
        item_price: parseFloat(document.getElementById(`fItemPrice-${idx}`).value || '0'),
      };
    });
  }

  window.addItemRowGlobal = addItemRow;     // ← เพิ่มบรรทัดนี้


  // หารค่าส่ง/ค่าอื่นๆ/ส่วนลดรวม ตามสัดส่วนราคาสินค้าแต่ละรายการ
  // เศษที่เหลือจากการปัดเศษ ยกให้รายการสุดท้าย กันผลรวมเพี้ยนจากราคาต้นทาง
  function allocateProportional(items, totalAmount){
    const subtotal = items.reduce((s, it) => s + (it.item_price || 0), 0);
    if (subtotal <= 0 || !totalAmount) return items.map(() => 0);
    const allocs = items.map(it => Math.round((totalAmount * (it.item_price / subtotal)) * 100) / 100);
    const allocatedSum = allocs.reduce((s, a) => s + a, 0);
    const diff = Math.round((totalAmount - allocatedSum) * 100) / 100;
    if (allocs.length > 0) allocs[allocs.length - 1] = Math.round((allocs[allocs.length - 1] + diff) * 100) / 100;
    return allocs;
  }

 function updateAllocPreview(){
  const items = getItemRows().filter(it => it.item_price > 0);
  const preview = document.getElementById('allocPreview');
  if (items.length === 0){ preview.innerHTML = ''; return; }

  const subtotal = items.reduce((s, it) => s + it.item_price, 0);
  const shipping = parseFloat(document.getElementById('fShipping').value || '0');
  const other = parseFloat(document.getElementById('fOther').value || '0');
  const discountTotal = collectDiscounts('discountList', subtotal).reduce((s, d) => s + Number(d.amount || 0), 0);

  const shipAlloc = allocateProportional(items, shipping);
  const otherAlloc = allocateProportional(items, other);
  const discAlloc = allocateProportional(items, discountTotal);

  const fmt = n => (Math.round(n * 100) / 100).toLocaleString('th-TH', { minimumFractionDigits: 0, maximumFractionDigits: 2 });

  let grandTotal = 0;
  const rows = items.map((it, i) => {
    const net = it.item_price + shipAlloc[i] + otherAlloc[i] - discAlloc[i];
    grandTotal += net;
    return `
      <div style="display:flex; justify-content:space-between; align-items:baseline; gap:10px; padding:8px 0; border-bottom:1px dashed var(--line-dash);">
        <div style="flex:1; min-width:0;">
          <div style="font-weight:600; color:var(--ink); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${it.product_group_name || 'สินค้า #' + (i+1)}</div>
          <div style="font-size:11px; color:var(--ink-soft); margin-top:2px;">
            ฿${fmt(it.item_price)} + ส่ง ฿${fmt(shipAlloc[i])} + อื่นๆ ฿${fmt(otherAlloc[i])} − ลด ฿${fmt(discAlloc[i])}
          </div>
        </div>
        <div style="flex-shrink:0; font-weight:700; color:var(--orange-dk); font-size:14px;">฿${fmt(net)}</div>
      </div>`;
  }).join('');

  preview.innerHTML = `
    <div style="font-weight:700; color:var(--ink); margin-bottom:6px;">📊 สรุปการหารสัดส่วนต้นทุนต่อชิ้น</div>
    ${rows}
    <div style="display:flex; justify-content:space-between; align-items:baseline; padding-top:10px; margin-top:2px;">
      <div style="font-weight:700; color:var(--ink);">รวมยอดสุทธิทั้งหมด</div>
      <div style="font-weight:800; color:var(--mint-dk); font-size:16px;">฿${fmt(grandTotal)}</div>
    </div>`;
}


  document.getElementById('submitBtn').onclick = async () => {
    const items = getItemRows();
    for (const it of items){
      if (!it.category_name || !it.product_group_name){ toast('กรุณากรอกหมวดสินค้า/กลุ่มสินค้าให้ครบทุกรายการ'); return; }
      if (!it.qty_buy || it.qty_buy <= 0){ toast('จำนวนซื้อต้องมากกว่า 0 ทุกรายการ'); return; }
      if (!it.ratio || it.ratio <= 0){ toast('เรโชต้องมากกว่า 0 ทุกรายการ'); return; }
      if (it.item_price == null || it.item_price < 0){ toast('ราคาสินค้าต้องไม่ติดลบ'); return; }
    }
    const storeName = document.getElementById('fStore').value.trim();
    if (!storeName){ toast('กรุณากรอกร้านค้า'); return; }

    const date = document.getElementById('fDate').value;
    const orderNo = document.getElementById('fOrderNo').value.trim();
    const recorder = currentEmail || (document.getElementById('fRecorder')?.value.trim() || '');
    const note = document.getElementById('fNote').value.trim();
    const shipping = parseFloat(document.getElementById('fShipping').value || '0');
    const other = parseFloat(document.getElementById('fOther').value || '0');
    const subtotal = items.reduce((s, it) => s + it.item_price, 0);
    const discountTotal = collectDiscounts('discountList', subtotal).reduce((s, d) => s + Number(d.amount || 0), 0);

    const shipAlloc = allocateProportional(items, shipping);
    const otherAlloc = allocateProportional(items, other);
    const discAlloc = allocateProportional(items, discountTotal);

    const btn = document.getElementById('submitBtn');
    btn.textContent = 'กำลังบันทึก...'; btn.disabled = true;

    try{
     // แก้ loop บันทึกในปุ่ม submit ของ renderAddPurchase()
    const receiptFile = document.getElementById('receiptInput').files[0];
      const billFile = document.getElementById('billInput').files[0];
      let firstProductGroupId = null;
      let receiptUrl = null;
      let billUrl = null;

      for (let i = 0; i < items.length; i++){
        const it = items[i];
        const payload = {
          date, category_name: it.category_name, product_group_name: it.product_group_name,
          product_url: it.product_url, store_name: storeName, order_no: orderNo,
          recorder, note,
          qty_buy: it.qty_buy, unit_buy: it.unit_buy, ratio: it.ratio, unit_sell: it.unit_sell,
          price_buy: it.item_price,
          shipping: shipAlloc[i],
          other_cost: otherAlloc[i],
          discounts: discAlloc[i] > 0 ? [{ name: 'ส่วนลดรวมออเดอร์ (หารตามสัดส่วน)', amount: discAlloc[i] }] : [],
        };
        const res = await gs('savePurchase', payload);
        if (i === 0) firstProductGroupId = res.product_group_id;

        // รูปใบเสร็จ/ใบวางบิล: อัปโหลดจริงแค่ครั้งแรก แล้วผูก URL เดียวกันซ้ำกับทุกสินค้าที่เหลือในออเดอร์นี้
        if (receiptFile){
          receiptUrl = receiptUrl
            ? await gs('setPurchaseImageUrl', res.id, 'receipt_img', receiptUrl).then(r => r.url)
            : await uploadPurchaseImageAndGetUrl(res.id, 'receipt_img', receiptFile);
        }
        if (billFile){
          billUrl = billUrl
            ? await gs('setPurchaseImageUrl', res.id, 'bill_img', billUrl).then(r => r.url)
            : await uploadPurchaseImageAndGetUrl(res.id, 'bill_img', billFile);
        }
      }

      toast(`บันทึกออเดอร์เรียบร้อย (${items.length} รายการ)`);
      go(`#/product/${firstProductGroupId}`);

    } finally {
      btn.textContent = '✔ บันทึกการซื้อทั้งออเดอร์'; btn.disabled = false;
    }
  };
}




function setupAutocomplete(inputId, listId, type, getCategoryId){
  const input = document.getElementById(inputId);
  const list = document.getElementById(listId);
  let debounce;

  async function showResults(kw){
    const results = await gs('searchMasterFor', type, kw, getCategoryId());
    list.innerHTML = '';
    const limit = kw ? 8 : 20; // ไม่พิมพ์คำค้น (แค่แตะเปิดดู) โชว์ได้เยอะกว่าปกติ
    results.slice(0, limit).forEach(r => {
      const item = el(`<div class="autocomplete-item">${r.name}</div>`);
      item.onclick = () => { input.value = r.name; list.style.display='none'; };
      list.appendChild(item);
    });
    if (kw && !results.find(r => r.name.toLowerCase() === kw.toLowerCase())){
      const item = el(`<div class="autocomplete-item create-new">+ เพิ่มใหม่ "${kw}"</div>`);
      item.onclick = () => { list.style.display='none'; };
      list.appendChild(item);
    }
    list.style.display = (results.length > 0 || kw) ? 'block' : 'none';
  }

  input.addEventListener('input', () => {
    clearTimeout(debounce);
    const kw = input.value.trim();
    debounce = setTimeout(() => showResults(kw), 250);
  });

  // แตะ/โฟกัสช่องตอนยังว่างอยู่ -> โชว์รายการทั้งหมดให้เลือกได้เลย ไม่ต้องพิมพ์ก่อน
  input.addEventListener('focus', () => {
    if (!input.value.trim()) showResults('');
  });

  input.addEventListener('blur', () => setTimeout(()=>list.style.display='none', 150));
}



//ตั้งราคาขาย
async function renderPricingPicker(){
  setTopbar({
    title: 'ตั้งราคาขาย', showBack: true,
    extra: `
      <div class="searchbox">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#3A2E28" stroke-width="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input id="ppSearch" placeholder="ค้นหากลุ่มสินค้า...">
      </div>`
  });
  content.innerHTML = `<div class="spinner"></div>`;

  let cards = await gs('getProductGroupCardsBatch', null);

  cards.sort((a, b) => a.name.localeCompare(b.name, 'th'));

  renderPPList(cards, '');
  document.getElementById('ppSearch').addEventListener('input', e => renderPPList(cards, e.target.value));
}

function renderPPList(cards, keyword){
  const kw = keyword.trim().toLowerCase();
  const filtered = kw ? cards.filter(c => c.name.toLowerCase().includes(kw)) : cards;
  content.innerHTML = '';
  if (filtered.length === 0){
    content.innerHTML = emptyState('💰', 'ไม่พบกลุ่มสินค้า', 'ลองคำค้นอื่น');
    return;
  }
  filtered.forEach(c => {
    const row = el(`
      <div class="row-card">
        <div class="avatar" style="background:var(--mint);">🏷️</div>
        <div class="row-main">
          <div class="row-title">${c.name}</div>
          <div class="row-sub">ต้นทุน/หน่วยขาย ${baht(c.cost_per_sell_unit)} · ขาย ${baht(c.sell_price_per_sell_unit)}</div>
        </div>
        <div class="row-arrow">›</div>
      </div>`);
    row.onclick = () => go(`#/pricing/${c.id}`);
    content.appendChild(row);
  });
}


async function renderStoreRanking(){
  setTopbar({ title: '🏆 ร้านแนะนำ', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const ranking = await gs('getStoreRanking');

  content.innerHTML = '';
  if (ranking.length === 0){
    content.innerHTML = emptyState('🏪', 'ยังไม่มีข้อมูลพอวิเคราะห์', 'บันทึกการซื้อสักพักแล้วกลับมาดูใหม่');
    return;
  }

  const cheapestWinner = ranking.slice().sort((a, b) => b.cheapest_count - a.cheapest_count)[0];
  const busiestWinner = ranking.slice().sort((a, b) => b.purchase_count - a.purchase_count)[0];
  const variedWinner = ranking.slice().sort((a, b) => b.product_variety - a.product_variety)[0];

  const crownCard = (icon, label, store) => `
    <div class="stat-tag" style="padding:12px;">
      <div class="tag-hole"></div>
      <div class="stat-label" style="font-size:11px;">${icon} ${label}</div>
      <div class="stat-value" style="font-size:15px;">${store.store_name}</div>
    </div>`;

  content.appendChild(el(`
  <div>
  <div class="grid">
    ${cheapestWinner && cheapestWinner.cheapest_count > 0 ? crownCard('🌱', 'ถูกที่สุด', cheapestWinner) : ''}
    ${busiestWinner ? crownCard('🔁', 'ซื้อบ่อยที่สุด', busiestWinner) : ''}
    ${variedWinner ? crownCard('🛍️', 'หลากหลายที่สุด', variedWinner) : ''}
  </div>
  <div class="section-title">📋 อันดับร้านค้าทั้งหมด</div>
  </div>
`));

  ranking.forEach((s, idx) => {
    const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `${idx + 1}.`;
    const row = el(`
      <div class="row-card">
        <div class="avatar">${medal}</div>
        <div class="row-main">
          <div class="row-title">${s.store_name}</div>
          <div class="row-sub">🌱 ถูกที่สุด ${s.cheapest_count} รายการ · 🔁 ซื้อ ${s.purchase_count} ครั้ง · 🛍️ ${s.product_variety} สินค้า</div>
          <div class="row-sub num" style="margin-top:2px;">ใช้จ่ายรวม ${baht(s.total_spend)}</div>
        </div>
      </div>`);
   content.appendChild(row);
  });

  const switchBtn = el(`<button class="btn btn-ghost" style="margin-top:16px;">💡สินค้าควรเปลี่ยนร้าน</button>`);
  switchBtn.onclick = () => go('#/switch-store');
  content.appendChild(switchBtn);
}


async function renderReorderList(){
  setTopbar({ title: '⏰ ควรซื้อเพิ่ม', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const cards = await gs('getProductGroupCardsBatch', null);
  const list = cards.filter(c => c.should_reorder);

  content.innerHTML = '';
  if (list.length === 0){
    content.innerHTML = emptyState('🎉', 'ไม่มีสินค้าที่ถึงรอบต้องซื้อเพิ่มตอนนี้', '');
    return;
  }

  content.appendChild(el(`<div style="font-size:13px; color:var(--ink-soft); margin-bottom:12px;">สินค้าที่ซื้อครั้งล่าสุดนานเกินรอบปกติแล้ว</div>`));

  list.forEach(c => {
    const row = el(`
      <div class="row-card">
        <div class="avatar" style="background:var(--yellow);">⏰</div>
        <div class="row-main">
          <div class="row-title">${c.name}</div>
          <div class="row-sub">🗓️ ซื้อล่าสุด ${c.last_purchase_date ? fmtDate(c.last_purchase_date) : '-'} · ปกติซื้อทุก ~${c.purchase_frequency_days} วัน (ผ่านมาแล้ว ${c.days_since_last_purchase} วัน)</div>
        </div>
        <div class="row-arrow">›</div>
      </div>`);
    row.onclick = () => go(`#/product/${c.id}`);
    content.appendChild(row);
  });
}





async function renderSwitchStore(){
  setTopbar({ title: '💡ควรเปลี่ยนร้าน', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const list = await gs('getSwitchStoreSuggestions');

  content.innerHTML = '';
  if (list.length === 0){
    content.innerHTML = emptyState('🎉', 'ซื้อร้านถูกสุดอยู่แล้วทุกรายการ', 'ไม่มีสินค้าที่แนะนำให้เปลี่ยนร้านตอนนี้');
    return;
  }

  content.appendChild(el(`<div style="font-size:13px; color:var(--ink-soft); margin-bottom:12px;">ซื้อครั้งล่าสุดจากร้านที่ไม่ใช่ร้านถูกสุด — เรียงจากประหยัดได้มากสุด</div>`));

  list.forEach(s => {
    const row = el(`
    <div>
      <div class="row-card">
        <div class="avatar" style="background:var(--yellow);">💡</div>
        <div class="row-main">
          <div class="row-title">${s.product_name}</div>
          <div class="row-sub">ซื้อจาก ${s.current_store_name} ${baht(s.current_price)}/${s.unit_sell}</div>
          <div class="row-sub" style="color:var(--mint-dk); font-weight:600;">🌱 ${s.cheapest_store_name} ${baht(s.cheapest_price)}/${s.unit_sell} — ประหยัด ${baht(s.savings_per_unit)}/${s.unit_sell}</div>
        </div>
        <div class="row-arrow">›</div>
      </div>
      </div>`);
    row.onclick = () => go(`#/product/${s.product_group_id}`);
    content.appendChild(row);
  });
}

/* ============================================================
   📊 Dashboard รายเดือน
   ============================================================ */
let dashboardState = { year: null, month: null }; // null = เดือนปัจจุบัน

async function renderDashboard(){
  const now = new Date();
  const year = dashboardState.year || now.getFullYear();
  const month = dashboardState.month || (now.getMonth() + 1);

  setTopbar({ title: '📊 Dashboard', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const d = await gs('getDashboardData', year, month);
  const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

  content.innerHTML = '';

  const isCurrentMonth = (year === now.getFullYear() && month === (now.getMonth() + 1));
  const monthNav = el(`
    <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:14px;">
      <button id="prevMonthBtn" style="padding:8px 16px; border-radius:999px; border:1.5px solid var(--line); background:none; font-size:16px;">‹</button>
      <div style="font-weight:700; font-size:16px;">${monthNames[month-1]} ${year + 543}</div>
      <button id="nextMonthBtn" style="padding:8px 16px; border-radius:999px; border:1.5px solid var(--line); background:none; font-size:16px; ${isCurrentMonth ? 'opacity:.3;' : ''}" ${isCurrentMonth ? 'disabled' : ''}>›</button>
    </div>`);
  content.appendChild(monthNav);

  document.getElementById('prevMonthBtn').onclick = () => {
    let m = month - 1, y = year;
    if (m < 1){ m = 12; y -= 1; }
    dashboardState = { year: y, month: m };
    renderDashboard();
  };
  if (!isCurrentMonth){
    document.getElementById('nextMonthBtn').onclick = () => {
      let m = month + 1, y = year;
      if (m > 12){ m = 1; y += 1; }
      dashboardState = { year: y, month: m };
      renderDashboard();
    };
  }

  if (d.purchase_count === 0){
    content.appendChild(el(emptyState('📊', 'ไม่มีข้อมูลเดือนนี้', 'ลองเลื่อนดูเดือนอื่น หรือบันทึกการซื้อใหม่')));
    return;
  }

  // ต่อจากนี้ใช้โค้ดแสดงผลเดิมทั้งหมดได้เลย — ลบแค่บรรทัดเดิมที่เคยแสดงชื่อเดือน/ปี
  // (บรรทัด content.appendChild(el(`<div style="font-weight:700...">${monthNames[now.getMonth()]}...`)) เดิม เอาออก เพราะย้ายมารวมกับ monthNav ด้านบนแล้ว)

  content.appendChild(el(`
    <div>
    <div class="grid">
      <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ใช้จ่ายรวม</div><div class="stat-value num">${baht(d.total_spend)}</div></div>
      <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">ซื้อทั้งหมด</div><div class="stat-value num">${d.purchase_count} ครั้ง · ${d.store_count} ร้าน</div></div>
      <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">🌱 ประหยัดได้</div><div class="stat-value num" style="color:var(--mint-dk);">${baht(d.total_saved)}</div></div>
      <div class="stat-tag"><div class="tag-hole"></div><div class="stat-label">💲 ปรับราคาขาย</div><div class="stat-value num">${d.price_drop_count > 0 ? `↓${d.price_drop_count}` : ''}${d.price_drop_count > 0 && d.price_rise_count > 0 ? ' · ' : ''}${d.price_rise_count > 0 ? `↑${d.price_rise_count}` : ''}${d.price_drop_count === 0 && d.price_rise_count === 0 ? '-' : ''}</div></div>
    </div>

    <div class="section-title">📉 ดัชนีต้นทุนโดยรวม</div>
    <div class="status-banner ${d.index_change_percent === null ? '' : d.index_change_percent <= -2 ? 'status-green' : d.index_change_percent >= 2 ? 'status-red' : 'status-yellow'}">
      <span style="font-size:20px;">${d.index_change_percent === null ? '📊' : d.index_change_percent <= -2 ? '↘' : d.index_change_percent >= 2 ? '↗' : '→'}</span>
      <div>${d.index_change_percent === null ? 'ข้อมูลยังไม่พอเทียบแนวโน้ม (ต้องมีประวัติซื้อข้ามหลายเดือน)' : `ต้นทุนรวม ${d.index_change_percent > 0 ? 'เพิ่มขึ้น' : d.index_change_percent < 0 ? 'ลดลง' : 'ทรงตัว'} ${Math.abs(d.index_change_percent)}% เทียบ 3 เดือนก่อน`}</div>
    </div>

    <div class="section-title">🤖 สรุปประจำเดือน</div>
    </div>
  `));

  const insightBlock = el(`<div class="info-block"></div>`);
  d.insights.forEach(txt => {
    insightBlock.appendChild(el(`<div style="font-size:14px; padding:6px 0; border-bottom:1px dashed var(--line);">${txt}</div>`));
  });
  content.appendChild(insightBlock);

  const linkRow = el(`<div style="display:flex; gap:8px; margin-top:14px;"></div>`);
  if (d.reorder_count > 0){
    const btn = el(`<button class="btn btn-ghost" style="flex:1;">⏰ควรซื้อเพิ่ม (${d.reorder_count})</button>`);
    btn.onclick = () => go('#/reorder-list');   // ← เปลี่ยนจาก '#/home'
      linkRow.appendChild(btn);
  }
  if (d.switch_suggestion_count > 0){
    const btn = el(`<button class="btn btn-ghost" style="flex:1;">💡ควรเปลี่ยนร้าน (${d.switch_suggestion_count})</button>`);
    btn.onclick = () => go('#/switch-store');
    linkRow.appendChild(btn);
  }
  if (linkRow.children.length > 0) content.appendChild(linkRow);
}


/* ============================================================
   ⚙️ จัดการข้อมูล — แก้ไข / ปิดใช้งาน / ลบ / รวม
   ============================================================ */
let masterState = { type: 'category', selectMode: false, selectedIds: new Set() };

const MASTER_TYPES = [
  { key:'category',     label:'หมวดสินค้าหลัก' },
  { key:'productGroup', label:'กลุ่มสินค้า' },
  { key:'store',        label:'ร้านค้า' },
  { key:'unit_buy',     label:'หน่วยซื้อ' },
  { key:'unit_sell',    label:'หน่วยขาย' },
  { key:'users',        label:'👤 ผู้ใช้งาน' }
];



async function renderMaster(){

  document.querySelectorAll('.select-bar, .select-fab').forEach(b => b.remove()); // ← เพิ่มบรรทัดนี้


  setTopbar({ title:'จัดการข้อมูล', showBack:true, extra:`<div class="tabs" id="masterTabs"></div>` });

  const tabsEl = document.getElementById('masterTabs');
  MASTER_TYPES.forEach(t => {
    const btn = el(`<button class="tab ${masterState.type===t.key?'active':''}">${t.label}</button>`);
    btn.onclick = () => {
      masterState.type = t.key; masterState.selectMode = false; masterState.selectedIds = new Set();
      masterState.categoryFilter = null; masterState.categories = null;
      renderMaster();
    };
    tabsEl.appendChild(btn);
  });

  content.innerHTML = `<div class="spinner"></div>`;
  if (masterState.type === 'productGroup' && !masterState.categories){
    masterState.categories = await gs('getCategories');
  }
  await refreshMasterList();
}

async function refreshMasterList(){
  document.querySelectorAll('.select-bar, .select-fab').forEach(b => b.remove());

  if (masterState.type === 'users'){ await renderUsersList(); return; }

  content.innerHTML = `<div class="spinner"></div>`; // โชว์สถานะโหลดทุกครั้ง กันหน้าดูเหมือนค้าง

  const isProductGroup = masterState.type === 'productGroup';
  const list = isProductGroup
    ? await gs('getMasterList', 'productGroup', masterState.categoryFilter)
    : await gs('getMasterList', masterState.type);

  content.innerHTML = '';

  // 🗂️ ข้อ 1: แท็บกลุ่มสินค้า — ให้เลือกหมวดก่อน ค่อยเลือกสินค้าที่จะย้าย/รวม
  if (isProductGroup){
    const catFilterWrap = el(`<div class="tabs" style="margin-bottom:10px;"></div>`);
    catFilterWrap.appendChild(el(`<button class="tab ${!masterState.categoryFilter ? 'active' : ''}" data-id="">ทั้งหมด</button>`));
    (masterState.categories || []).forEach(c => {
      catFilterWrap.appendChild(el(`<button class="tab ${masterState.categoryFilter===c.id?'active':''}" data-id="${c.id}">${c.name}</button>`));
    });
    catFilterWrap.querySelectorAll('.tab').forEach(btn => {
      btn.onclick = () => {
        masterState.categoryFilter = btn.dataset.id || null;
        masterState.selectMode = false; masterState.selectedIds = new Set();
        refreshMasterList();
      };
    });
    content.appendChild(catFilterWrap);
  }

  const selectLabel = masterState.selectMode
    ? '✕ ยกเลิกเลือก'
    : (isProductGroup ? '🔗 เลือกเพื่อรวม/ย้ายหมวด' : '🔗 เลือกเพื่อรวมรายการซ้ำ');

  const actionRow = el(`<div style="display:flex; justify-content:${masterState.type === 'store' ? 'space-between' : 'flex-end'}; margin-bottom:10px; gap:8px;">
    ${masterState.type === 'store' ? `<button class="btn btn-ghost" style="width:auto; padding:8px 16px; font-size:13px;" id="storeRankingBtn">🏆 ดูอันดับร้าน</button>` : ''}
    <button class="btn btn-ghost" style="width:auto; padding:8px 16px; font-size:13px;" id="toggleSelectBtn">${selectLabel}</button>
  </div>`);
  content.appendChild(actionRow);

  document.getElementById('toggleSelectBtn').onclick = () => {
    masterState.selectMode = !masterState.selectMode;
    masterState.selectedIds = new Set();
    refreshMasterList();
  };

  const rankBtn = document.getElementById('storeRankingBtn');
  if (rankBtn) rankBtn.onclick = () => go('#/store-ranking');

  if (masterState.type === 'category' && !masterState.selectMode){
    const addBlock = el(`
      <div class="info-block" style="margin-bottom:14px;">
        <div class="info-block-title">➕ เพิ่มหมวดสินค้า</div>
        <div class="field" style="display:flex; gap:8px;">
          <input id="newCategoryName" placeholder="ชื่อหมวดสินค้า เช่น เครื่องเขียน" style="flex:1;">
          <button class="btn btn-primary" id="addCategoryBtn" style="width:auto; padding:0 18px;">เพิ่ม</button>
        </div>
      </div>`);
    content.appendChild(addBlock);

    const nameInput = document.getElementById('newCategoryName');
    const doAdd = async () => {
      const name = nameInput.value.trim();
      if (!name){ toast('กรุณากรอกชื่อหมวดสินค้า'); return; }
      const btn = document.getElementById('addCategoryBtn');
      btn.disabled = true;
      try{
        await gs('addMaster', 'category', name);
        toast('เพิ่มหมวดสินค้าแล้ว');
        masterState.categories = null; // เพิ่มหมวดใหม่แล้ว ล้าง cache ให้แท็บกลุ่มสินค้าดึงใหม่
        refreshMasterList();
      } catch(e){ toast(e.message || 'เพิ่มไม่สำเร็จ'); }
      finally { btn.disabled = false; }
    };
    document.getElementById('addCategoryBtn').onclick = doAdd;
    nameInput.addEventListener('keydown', ev => { if (ev.key === 'Enter') doAdd(); });
  }

  if (list.length === 0){
    content.appendChild(el(emptyState('📁', isProductGroup && masterState.categoryFilter ? 'ไม่มีกลุ่มสินค้าในหมวดนี้' : 'ยังไม่มีข้อมูล', 'กด + บันทึกการซื้อ เพื่อเริ่มสร้างข้อมูล')));
    return;
  }

  list.forEach(item => content.appendChild(masterRow(item)));

   if (masterState.selectMode && masterState.selectedIds.size >= 1){
    const canMerge = masterState.selectedIds.size >= 2;
    const fab = el(`
      <button class="select-fab">
        📋 จัดการที่เลือก <span class="count-badge">${masterState.selectedIds.size}</span>
      </button>`);
    fab.onclick = () => openSelectSheet(isProductGroup, canMerge);
    document.body.appendChild(fab);
  }
}

function openSelectSheet(isProductGroup, canMerge){
  const overlay = el(`
    <div class="select-sheet-overlay">
      <div class="select-sheet">
        <div class="select-sheet-header">
          <span>เลือกแล้ว ${masterState.selectedIds.size} รายการ</span>
          <button class="close-x">✕</button>
        </div>
        <div class="select-sheet-actions">
          ${isProductGroup ? `<button class="primary-move">📂 ย้ายไปหมวดอื่น</button>` : ''}
          ${canMerge ? `<button class="primary-merge">รวมรายการ</button>` : ''}
          <button class="cancel-select">ยกเลิกการเลือกทั้งหมด</button>
        </div>
      </div>
    </div>`);

  overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
  overlay.querySelector('.close-x').onclick = () => overlay.remove();

  const moveBtn = overlay.querySelector('.primary-move');
  if (moveBtn) moveBtn.onclick = () => {
    overlay.remove();
    renderMovePicker().catch(err => { console.error(err); toast('เปิดหน้าย้ายหมวดไม่สำเร็จ: ' + (err.message||err)); });
  };
  const mergeBtn = overlay.querySelector('.primary-merge');
  if (mergeBtn) mergeBtn.onclick = () => {
    overlay.remove();
    renderMergePicker().catch(err => { console.error(err); toast('เปิดหน้ารวมรายการไม่สำเร็จ: ' + (err.message||err)); });
  };
  overlay.querySelector('.cancel-select').onclick = () => {
    masterState.selectMode = false; masterState.selectedIds = new Set();
    overlay.remove();
    refreshMasterList();
  };

  document.body.appendChild(overlay);
}



/* ============================================================
   📂 ย้ายกลุ่มสินค้าไปหมวดอื่นทีละหลายรายการ
   วางต่อท้ายฟังก์ชัน renderMergePicker() เดิม
   ============================================================ */
async function renderMovePicker(){
  document.querySelectorAll('.select-bar').forEach(b => b.remove());
  setTopbar({ title:'ย้ายไปหมวดสินค้าอื่น', showBack:true });
  content.innerHTML = `<div class="spinner"></div>`; // ← เพิ่มบรรทัดนี้
  const ids = Array.from(masterState.selectedIds);
  const categories = await gs('getMasterList', 'category');

  content.innerHTML = `<div style="margin-bottom:14px; color:var(--ink-soft); font-size:14px;">
    เลือกหมวดสินค้าปลายทาง — กลุ่มสินค้าที่เลือกไว้ ${ids.length} รายการ จะถูกย้ายไปหมวดนี้ทั้งหมด
    (ถ้ารายการไหนชื่อชนกับที่มีอยู่แล้วในหมวดปลายทาง จะข้ามรายการนั้นให้อัตโนมัติ)
  </div>`;

  let picked = categories.length > 0 ? categories[0].id : null;
  const listWrap = el(`<div></div>`);
  content.appendChild(listWrap);

  function draw(){
    listWrap.innerHTML = '';
    categories.forEach(c => {
      const row = el(`
        <div class="merge-target-row ${picked === c.id ? 'picked' : ''}">
          <div class="radio-dot ${picked === c.id ? 'on' : ''}"></div>
          <div style="flex:1;"><div style="font-weight:700;">${c.name}</div><div class="row-sub">มีกลุ่มสินค้าอยู่แล้ว ${c.usage} รายการ</div></div>
        </div>`);
      row.onclick = () => { picked = c.id; draw(); };
      listWrap.appendChild(row);
    });
  }
  draw();

 const confirmBtn = el(`<button class="btn btn-mint" style="margin-top:10px;">✔ ย้ายรายการ</button>`);
const confirmBtnLabel = confirmBtn.textContent;
confirmBtn.onclick = async () => {
  if (!picked){ toast('กรุณาเลือกหมวดปลายทาง'); return; }
  confirmBtn.disabled = true;
  confirmBtn.innerHTML = `<span class="btn-spinner"></span> กำลังย้าย...`;
  try{
    const res = await gs('moveProductGroupsToCategory', ids, picked);
    toast(res.skipped > 0
      ? `ย้ายแล้ว ${res.moved} รายการ (ข้าม ${res.skipped} รายการที่ชื่อซ้ำ)`
      : `ย้ายแล้ว ${res.moved} รายการ`);
    masterState.selectMode = false; masterState.selectedIds = new Set();
    renderMaster();
  } catch(err){
    console.error('moveProductGroupsToCategory error:', err);
    toast('ย้ายไม่สำเร็จ: ' + (err.message || err));
  } finally {
    confirmBtn.textContent = confirmBtnLabel; confirmBtn.disabled = false;
  }
};
content.appendChild(confirmBtn);


}

const BTN_SPINNER_CSS = `
<style>
.btn-spinner{
  display:inline-block; width:14px; height:14px; margin-right:6px;
  border:2px solid rgba(255,255,255,.5); border-top-color:#fff;
  border-radius:50%; vertical-align:-2px; animation:btnspin .7s linear infinite;
}
@keyframes btnspin{ to{ transform:rotate(360deg); } }
</style>`;



if (!document.getElementById('btnSpinnerStyleTag')){
  const tag = document.createElement('div');
  tag.id = 'btnSpinnerStyleTag';
  tag.innerHTML = BTN_SPINNER_CSS;
  document.head.appendChild(tag.firstElementChild);
}


function masterRow(item){
  const isSelected = masterState.selectedIds.has(item.id);
  const row = el(`<div class="master-row ${!item.active ? 'inactive' : ''} ${isSelected ? 'selected' : ''}"></div>`);

  if (masterState.selectMode){
    row.appendChild(el(`<div class="master-check ${isSelected ? 'checked' : ''}">${isSelected ? '✓' : ''}</div>`));
    row.appendChild(el(`<div class="master-name">${item.name}</div>`));
    row.appendChild(el(`<div class="master-usage">ใช้งาน ${item.usage}</div>`));
    row.onclick = () => {
      if (masterState.selectedIds.has(item.id)) masterState.selectedIds.delete(item.id);
      else masterState.selectedIds.add(item.id);
      refreshMasterList();
    };
    return row;
  }

  const nameWrap = el(`<div class="master-name">${item.name}</div>`);
  row.appendChild(nameWrap);
  row.appendChild(el(`<div class="master-usage">ใช้งาน ${item.usage}</div>`));

  const editBtn = el(`<button class="icon-btn">✏️</button>`);
  editBtn.onclick = (e) => {
    e.stopPropagation();
    nameWrap.innerHTML = `<input value="${item.name}">`;
    const input = nameWrap.querySelector('input');
    input.focus(); input.select();
    let done = false;
    const commit = async () => {
      if (done) return; done = true;
      const newName = input.value.trim();
      if (!newName || newName === item.name){ refreshMasterList(); return; }
      try{
        await gs('renameMaster', masterState.type, item.id, newName);
        toast('บันทึกแล้ว');
      } finally { refreshMasterList(); }
    };
    input.addEventListener('keydown', ev => {
      if (ev.key === 'Enter') commit();
      if (ev.key === 'Escape'){ done = true; refreshMasterList(); }
    });
    input.addEventListener('blur', commit);
  };
  row.appendChild(editBtn);

  const isUnit = masterState.type === 'unit_buy' || masterState.type === 'unit_sell';
  if (!isUnit){
    const sw = el(`<button class="switch ${item.active ? 'on' : ''}"></button>`);
    sw.onclick = async (e) => {
      e.stopPropagation();
      try{
        await gs('setMasterActive', masterState.type, item.id, !item.active);
        toast(item.active ? 'ปิดใช้งานแล้ว' : 'เปิดใช้งานแล้ว');
      } finally { refreshMasterList(); }
    };
    row.appendChild(sw);
  }

  const delBtn = el(`<button class="icon-btn danger" ${item.usage > 0 ? 'disabled' : ''}>🗑️</button>`);
  delBtn.onclick = async (e) => {
    e.stopPropagation();
    if (item.usage > 0){ toast('มีการใช้งานอยู่ ลบไม่ได้'); return; }
    if (!(await customConfirm(`ลบ "${item.name}" ถาวร?`))) return;
    try{
      await gs('deleteMasterPermanent', masterState.type, item.id);
      toast('ลบแล้ว');
    } finally { refreshMasterList(); }
  };
  row.appendChild(delBtn);

  return row;
}

async function renderUsersList(){
  const users = await gs('getAllowedUsers');
  content.innerHTML = '';

  content.appendChild(el(`
    <div class="info-block">
      <div class="info-block-title">➕ เพิ่มผู้ใช้งาน</div>
      <div class="field"><label>อีเมล Google</label><input id="newUserEmail" placeholder="name@gmail.com"></div>
      <div class="field"><label>ชื่อ (ไม่บังคับ)</label><input id="newUserName" placeholder="ชื่อผู้ใช้"></div>
      <button class="btn btn-primary" id="addUserBtn">เพิ่มผู้ใช้งาน</button>
    </div>
  `));
  document.getElementById('addUserBtn').onclick = async () => {
    const email = document.getElementById('newUserEmail').value.trim();
    const name = document.getElementById('newUserName').value.trim();
    if (!email){ toast('กรุณากรอกอีเมล'); return; }
    const btn = document.getElementById('addUserBtn');
    btn.disabled = true;
    try{
      await gs('addAllowedUser', email, name);
      toast('เพิ่มผู้ใช้งานแล้ว');
      refreshMasterList();
    } finally { btn.disabled = false; }
  };

  content.appendChild(el(`<div class="section-title">👥 รายชื่อที่เข้าใช้งานได้ (${users.length} คน)</div>`));

  if (users.length === 0){
    content.appendChild(el(emptyState('👤','ยังไม่มีผู้ใช้งานในระบบ','')));
    return;
  }

  users.forEach(u => {
    const row = el(`
      <div class="master-row ${!u.active ? 'inactive' : ''}">
        <div class="master-name">${u.name ? u.name + ' — ' : ''}${u.email}</div>
      </div>`);
    const sw = el(`<button class="switch ${u.active ? 'on' : ''}"></button>`);
    sw.onclick = async () => {
      try{ await gs('setAllowedUserActive', u.row, !u.active); toast(u.active ? 'ปิดสิทธิ์แล้ว' : 'เปิดสิทธิ์แล้ว'); }
      finally { refreshMasterList(); }
    };
    const delBtn = el(`<button class="icon-btn danger">🗑️</button>`);
    delBtn.onclick = async () => {
    if (!(await customConfirm(`ลบสิทธิ์การเข้าใช้งานของ ${u.email}?`))) return;
      try{ await gs('removeAllowedUser', u.row); toast('ลบแล้ว'); }
      finally { refreshMasterList(); }
    };
    row.appendChild(sw);
    row.appendChild(delBtn);
    content.appendChild(row);
  });
}



async function renderMergePicker(){
  document.querySelectorAll('.select-bar').forEach(b => b.remove());
  setTopbar({ title:'เลือกรายการหลัก', showBack:true });
  content.innerHTML = `<div class="spinner"></div>`; // ← เพิ่มบรรทัดนี้
  const ids = Array.from(masterState.selectedIds);
  const list = await gs('getMasterList', masterState.type);
  const items = list.filter(i => ids.includes(i.id));

  content.innerHTML = `<div style="margin-bottom:14px; color:var(--ink-soft); font-size:14px;">
    เลือกรายการที่จะเก็บไว้ — รายการอื่นจะถูกรวมเข้ามา และข้อมูลที่เกี่ยวข้องทั้งหมดจะย้ายมาที่รายการนี้ (ลบรายการอื่นถาวร)
  </div>`;

  let picked = items[0].id;
  const listWrap = el(`<div></div>`);
  content.appendChild(listWrap);

  function draw(){
    listWrap.innerHTML = '';
    items.forEach(i => {
      const row = el(`
        <div class="merge-target-row ${picked === i.id ? 'picked' : ''}">
          <div class="radio-dot ${picked === i.id ? 'on' : ''}"></div>
          <div style="flex:1;"><div style="font-weight:700;">${i.name}</div><div class="row-sub">ใช้งาน ${i.usage} รายการ</div></div>
        </div>`);
      row.onclick = () => { picked = i.id; draw(); };
      listWrap.appendChild(row);
    });
  }
  draw();

  const confirmBtn = el(`<button class="btn btn-mint" style="margin-top:10px;">✔ รวมรายการ</button>`);
  confirmBtn.onclick = async () => {
    confirmBtn.textContent = 'กำลังรวม...'; confirmBtn.disabled = true;
    try{
      const sourceIds = ids.filter(id => id !== picked);
      await gs('mergeMaster', masterState.type, sourceIds, picked);
      toast('รวมรายการเรียบร้อย');
      masterState.selectMode = false; masterState.selectedIds = new Set();
      renderMaster();
    } finally { confirmBtn.textContent = '✔ รวมรายการ'; confirmBtn.disabled = false; }
  };
  content.appendChild(confirmBtn);
}


/* ============================================================
   📊 ประวัติราคา — เลือกกลุ่มสินค้า แล้วดูประวัติการเปลี่ยนราคา
   ============================================================ */
async function renderPriceHistoryPicker(){
  setTopbar({
    title: 'ประวัติเปลี่ยนราคาขาย', showBack: true,
    extra: `
      <div class="searchbox">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#3A2E28" stroke-width="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input id="phSearch" placeholder="ค้นหากลุ่มสินค้า...">
      </div>`
  });
  content.innerHTML = `<div class="spinner"></div>`;

 let cards = await gs('getProductGroupCardsBatch', null);
  cards.sort((a, b) => a.name.localeCompare(b.name, 'th'));

  renderPHList(cards, '');
  document.getElementById('phSearch').addEventListener('input', e => renderPHList(cards, e.target.value));
}

function renderPHList(cards, keyword){
  const kw = keyword.trim().toLowerCase();
  const filtered = kw ? cards.filter(c => c.name.toLowerCase().includes(kw)) : cards;
  content.innerHTML = '';
  if (filtered.length === 0){
    content.innerHTML = emptyState('📊', 'ไม่พบกลุ่มสินค้า', 'ลองคำค้นอื่น');
    return;
  }
  filtered.forEach(c => {
    const row = el(`
      <div class="row-card">
        <div class="avatar" style="background:var(--orange);">🏷️</div>
        <div class="row-main">
          <div class="row-title">${c.name}</div>
          <div class="row-sub">ราคาขายปัจจุบัน ${baht(c.sell_price)}</div>
        </div>
        <div class="row-arrow">›</div>
      </div>`);
    row.onclick = () => go(`#/price-history/${c.id}`);
    content.appendChild(row);
  });
}

async function renderPriceHistoryDetail(pgId){
  setTopbar({ title: 'ประวัติราคา', showBack: true });
  content.innerHTML = `<div class="spinner"></div>`;

  const d = await gs('getPriceHistoryDetail', pgId);
  const seasonal = await gs('getSeasonalPriceAnalysis', pgId).catch(() => null);

  content.innerHTML = '';
  content.appendChild(el(`
    <div class="tag-card"><div class="tag-hole"></div>
      <div class="tag-media">${d.image_url ? `<img src="${d.image_url}">` : '🏷️'}</div>
      <div class="tag-body"><div class="tag-name" style="font-size:19px;">${d.name}</div></div>
    </div>`));

  if (seasonal){
    const trendBadge = seasonal.trend
      ? (seasonal.trend.direction === 'up'
          ? `<span class="badge badge-red">↗ ต้นทุนเพิ่มขึ้น ${seasonal.trend.percent}%</span>`
          : seasonal.trend.direction === 'down'
            ? `<span class="badge badge-mint">↘ ต้นทุนลดลง ${seasonal.trend.percent}%</span>`
            : `<span class="badge" style="background:var(--line); color:var(--ink-soft);">→ ทรงตัว</span>`)
      : '';
    content.appendChild(el(`
      <div class="info-block">
        <div class="info-block-title">📅 ช่วงเวลาซื้อถูก/แพงที่สุด ${trendBadge}</div>
        <div class="info-row"><span class="k">ซื้อทั้งหมด</span><span class="v num">${seasonal.purchase_count} ครั้ง</span></div>
        <div class="info-row"><span class="k">ต้นทุนเฉลี่ย/หน่วยขาย</span><span class="v num">${baht(seasonal.avg_cost)}</span></div>
        ${seasonal.cheapest_month ? `<div class="info-row"><span class="k">🌱 ถูกสุดช่วง</span><span class="v num" style="color:var(--mint-dk);">${seasonal.cheapest_month.month_name} (${baht(seasonal.cheapest_month.avg)})</span></div>` : ''}
        ${seasonal.priciest_month ? `<div class="info-row"><span class="k">🔺 แพงสุดช่วง</span><span class="v num" style="color:var(--danger);">${seasonal.priciest_month.month_name} (${baht(seasonal.priciest_month.avg)})</span></div>` : ''}
        ${!seasonal.cheapest_month ? `<div style="font-size:12px; color:var(--ink-soft); margin-top:6px;">ต้องซื้ออย่างน้อย 2 เดือนต่างกันถึงจะเทียบช่วงเวลาได้</div>` : ''}
      </div>`));
  }

  if (d.history.length === 0){
    content.appendChild(el(emptyState('📊', 'ยังไม่มีประวัติการเปลี่ยนราคาขาย', 'ราคาขายของกลุ่มนี้ยังไม่เคยถูกปรับ')));
    return;
  }

  if (d.history.length >= 2){
    const chartCard = el(`<div class="info-block"><div class="info-block-title">📈 แนวโน้มราคาขาย</div><div id="sparkWrap"></div></div>`);
    content.appendChild(chartCard);
    document.getElementById('sparkWrap').innerHTML = renderSparkline(d.history);
  }

  content.appendChild(el(`<div class="section-title">📜 ประวัติการเปลี่ยนราคา (ใหม่ → เก่า)</div>`));
  d.history.forEach(h => content.appendChild(priceHistoryRow(h)));
}


function priceHistoryRow(h){
  const changed = Math.round((h.new_price - h.old_price) * 100) / 100;
  let badge;
  if (changed > 0) badge = `<span class="badge badge-red">▲ +${baht(changed)}</span>`;
  else if (changed < 0) badge = `<span class="badge badge-mint">▼ ${baht(Math.abs(changed))}</span>`;
  else badge = `<span class="badge" style="background:var(--line); color:var(--ink-soft);">คงเดิม</span>`;

  return el(`
    <div class="info-block" style="margin-bottom:10px;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
        <div style="font-weight:700;">${fmtDate(h.date)}</div>
        ${badge}
      </div>
      <div class="info-row"><span class="k">ราคาขาย</span><span class="v num">${baht(h.old_price)} → ${baht(h.new_price)}</span></div>
      <div class="info-row"><span class="k">ต้นทุนตอนนั้น</span><span class="v num">${baht(h.cost_price)}</span></div>
      <div class="info-row"><span class="k">กำไร/หน่วย</span><span class="v num">${baht(h.profit)}</span></div>
      <div class="info-row"><span class="k">Margin</span><span class="v num">${h.margin}%</span></div>
      ${h.edited_by ? `<div class="info-row"><span class="k">ผู้แก้ไข</span><span class="v">${h.edited_by}</span></div>` : ''}
    </div>`);
}

//เพิ่มฟังก์ชันใหม่ทั้งหมด — วางไว้ก่อนส่วน `Dialog กำหนดเอง

let _webAppUrlCache = null;

async function showProductQR(pgId, productName){
  const overlay = el(`
    <div class="modal-overlay">
      <div class="modal-box" style="max-width:320px; text-align:center;">
        <div class="modal-message" style="margin-bottom:4px; font-weight:700;">${productName}</div>
        <div style="font-size:12px; color:var(--ink-soft); margin-bottom:16px;">สแกนเพื่อเปิดหน้าสินค้านี้โดยตรง</div>
        <div id="qrCanvasWrap" style="display:flex; justify-content:center; margin-bottom:16px; min-height:200px; align-items:center;">
          <div class="spinner" style="margin:0;"></div>
        </div>
        <div id="qrLinkText" style="font-size:11px; color:var(--ink-soft); word-break:break-all; margin-bottom:16px;"></div>
        <div class="modal-actions">
          <button class="btn btn-ghost modal-cancel">ปิด</button>
          <button class="btn btn-primary qr-copy-btn">คัดลอกลิงก์</button>
        </div>
      </div>
    </div>`);
  overlay.querySelector('.modal-cancel').onclick = () => overlay.remove();
  document.body.appendChild(overlay);

  try{
    if (!_webAppUrlCache) _webAppUrlCache = await gs('getWebAppUrl');
    
    const link = _webAppUrlCache + '?page=product&pgId=' + pgId;

    const wrap = overlay.querySelector('#qrCanvasWrap');
    wrap.innerHTML = '';
    new QRCode(wrap, { text: link, width: 200, height: 200, colorDark: '#3A2E28', colorLight: '#FFFFFF' });
    overlay.querySelector('#qrLinkText').textContent = link;

    overlay.querySelector('.qr-copy-btn').onclick = async () => {
      try{
        await navigator.clipboard.writeText(link);
        toast('คัดลอกลิงก์แล้ว');
      } catch(err){
        toast('คัดลอกไม่สำเร็จ ลองกดค้างที่ลิงก์ด้านล่างแทน');
      }
    };
  } catch(err){
    overlay.querySelector('#qrCanvasWrap').innerHTML = `<div style="color:var(--danger); font-size:13px;">สร้าง QR ไม่สำเร็จ: ${err.message || err}</div>`;
  }
}






/** เส้นกราฟราคาขายง่ายๆ ด้วย SVG (ไม่พึ่ง library ภายนอก) */
function renderSparkline(history){
  const points = history.slice().reverse().map(h => ({ date: new Date(h.date), price: h.new_price }));
  if (points.length < 2) return '';

  const w = 320, h = 84, pad = 10;
  const prices = points.map(p => p.price);
  const min = Math.min(...prices), max = Math.max(...prices);
  const range = (max - min) || 1;
  const stepX = (w - 2 * pad) / (points.length - 1);

  const coords = points.map((p, i) => {
    const x = pad + i * stepX;
    const y = h - pad - ((p.price - min) / range) * (h - 2 * pad);
    return [x.toFixed(1), y.toFixed(1)];
  });
  const polyStr = coords.map(c => c.join(',')).join(' ');
  const dots = coords.map(([x, y]) => `<circle cx="${x}" cy="${y}" r="3.2" fill="#FF6B4A"/>`).join('');

  return `<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="display:block;">
    <polyline points="${polyStr}" fill="none" stroke="#FF6B4A" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
    ${dots}
  </svg>
  <div style="display:flex; justify-content:space-between; font-size:11.5px; color:var(--ink-soft); margin-top:4px;">
    <span>${fmtDate(points[0].date)}</span><span>${fmtDate(points[points.length-1].date)}</span>
  </div>`;
}


// ======================================
// GitHub QR Scanner
// ======================================

const GITHUB_ORIGIN = "https://kunjay-eng.github.io";
const SCANNER_URL =
  "https://kunjay-eng.github.io/product-manager/scanner.html";

//const OCR_READER_URL = "https://kunjay-eng.github.io/product-manager/ocr-reader.html";


let qrSession = null;
let scannerWindow = null;

// ------------------------------
// เปิด Scanner
// ------------------------------
function openQRScannerExternal() {

  qrSession = crypto.randomUUID();

  scannerWindow = window.open(

    SCANNER_URL +
    "?session=" +
    encodeURIComponent(qrSession),

    "qrScanner",

    "width=420,height=760"

  );

}

// ------------------------------
// รับค่าจาก GitHub
// ------------------------------
window.addEventListener("message", function (event) {

  // ตรวจว่าเป็น GitHub ของเราหรือไม่
  if (event.origin !== GITHUB_ORIGIN)
    return;

  const data = event.data;

  if (!data)
    return;

  if (data.type !== "QR_RESULT")
    return;

  if (data.session !== qrSession)
    return;

  // ปิดหน้าสแกน (ถ้ายังไม่ปิด)
  try {
    if (scannerWindow && !scannerWindow.closed) {
      scannerWindow.close();
    }
  } catch (e) {}

  qrSession = null;

  handleScannedQR(data.text);

});



function handleScannedQR(text){
  let pgId = null;
  try{
    const url = new URL(text);
    if (url.searchParams.get('page') === 'product' && url.searchParams.get('pgId')){
      pgId = url.searchParams.get('pgId');
    } else if (url.hash.includes('#/product/')){
      pgId = url.hash.split('#/product/')[1].split(/[?&#]/)[0];
    }
  } catch(err){
    // text ที่สแกนได้ไม่ใช่ URL ที่ parse ได้ (เช่น QR ทั่วไปที่ไม่ใช่ลิงก์สินค้า) -> ปล่อย pgId เป็น null
  }

  if (pgId){
    toast('พบสินค้า กำลังเปิด...');
    go(`#/product/${pgId}`);
  } else {
    toast('QR นี้ไม่ใช่ลิงก์สินค้าในระบบนี้');
  }
}

// ======================================
// GitHub OCR Reader (อ่านรูปคำสั่งซื้อด้วย AI)
// ======================================
const OCR_READER_URL = "https://kunjay-eng.github.io/product-manager/ocr-reader.html";

let ocrSession = null;
let ocrWindow = null;
let ocrInitData = null;       // แคชหมวด/กลุ่มสินค้า กันดึงซ้ำทุกครั้งที่เปิด
let ocrReadyReceived = false; // true เมื่อหน้าต่าง OCR แจ้งว่าพร้อมรับข้อมูลแล้ว

function openOCRReaderExternal(){
  ocrSession = crypto.randomUUID();
  ocrReadyReceived = false;

  // ต้องเปิดหน้าต่างแบบ synchronous ก่อน await ใดๆ เสมอ กัน popup blocker
  ocrWindow = window.open(
    OCR_READER_URL + "?session=" + encodeURIComponent(ocrSession),
    "ocrReader",
    "width=460,height=800"
  );

  (async () => {
    if (!ocrInitData){
      try{ ocrInitData = await gs('getMasterDataForOCR'); }
      catch(err){ console.error('โหลดข้อมูลหมวด/กลุ่มสินค้าไม่สำเร็จ:', err); ocrInitData = { categories: [], productGroups: [] }; }
    }
    sendOcrInitDataIfReady();
  })();
}

function sendOcrInitDataIfReady(){
  if (ocrReadyReceived && ocrInitData && ocrWindow && !ocrWindow.closed){
    ocrWindow.postMessage({ type: "INIT_DATA", session: ocrSession, payload: ocrInitData }, GITHUB_ORIGIN);
  }
}

window.addEventListener("message", function (event) {
  if (event.origin !== GITHUB_ORIGIN) return;
  const data = event.data;
  if (!data) return;

  if (data.type === "OCR_READY" && data.session === ocrSession){
    ocrReadyReceived = true;
    sendOcrInitDataIfReady();
    return;
  }

  if (data.type !== "OCR_RESULT") return;
  if (data.session !== ocrSession) return;

  try { if (ocrWindow && !ocrWindow.closed) ocrWindow.close(); } catch (e) {}
  ocrSession = null;

  handleOCRResult(data.data);
});


/**
 * data = {
 *   store_name, order_date, order_no, shipping, other_cost,
 *   discounts: [{name, amount}],
 *   items: [{category_name, product_name, product_url, qty, unit_buy, ratio, unit_sell, unit_price, total_price}]
 * }
 * เติมข้อมูลลงในฟอร์ม "บันทึกการซื้อ" ที่เปิดอยู่ (renderAddPurchase) แบบ 1:1 ทุกช่อง
 * ยังไม่บันทึกลง Sheet — ผู้ใช้ต้องตรวจสอบแล้วกด "✔ บันทึกการซื้อทั้งออเดอร์" เองเหมือนเดิม
 */
function handleOCRResult(data){
  if (!data || !document.getElementById('orderItemsList')) {
    toast('ไม่พบฟอร์มบันทึกการซื้อ กรุณาเปิดหน้า "บันทึกซื้อ" แล้วลองใหม่');
    return;
  }

  // หัวออเดอร์
  if (data.store_name) document.getElementById('fStore').value = data.store_name;
  if (data.order_no)   document.getElementById('fOrderNo').value = data.order_no;
  if (data.order_date && /^\d{4}-\d{2}-\d{2}$/.test(data.order_date)) {
    document.getElementById('fDate').value = data.order_date;
  }

  // ค่าขนส่ง / ค่าอื่นๆ
  if (data.shipping)   document.getElementById('fShipping').value = data.shipping;
  if (data.other_cost) document.getElementById('fOther').value = data.other_cost;

  // ส่วนลด — เติมลงรายการส่วนลดที่มีอยู่แล้ว (ใช้ discountRowEl เดิมของระบบ)
  if (Array.isArray(data.discounts) && data.discounts.length > 0){
    const discList = document.getElementById('discountList');
    data.discounts.forEach(d => {
      if (d && Number(d.amount) > 0) discList.appendChild(discountRowEl(d.name || '', d.amount, 'amount'));
    });
  }

  // ล้างรายการสินค้าว่างที่ auto-add ไว้ตอนเปิดหน้า โดย "ใช้ซ้ำ" แถวนั้นเป็นสินค้ารายการแรก
  // (ห้ามลบทิ้งแล้วสร้างใหม่ เพราะ orderItemSeq เป็นตัวนับเดินหน้าอย่างเดียว ไม่ลดค่าเมื่อลบแถว
  //  ถ้าลบแล้วสร้างใหม่ เลข "สินค้าที่ N" จะกระโดดเริ่มที่ 2 ทันที)
  const itemsList = document.getElementById('orderItemsList');
  const existingRows = Array.from(itemsList.querySelectorAll('.order-item'));
  const firstRowEmpty = existingRows.length === 1 &&
    !document.getElementById(`fGroup-${existingRows[0].dataset.idx}`).value.trim();

  const ocrItems = data.items || [];
  let usedFirstRow = false;

  if (firstRowEmpty && ocrItems.length > 0){
    fillItemFields(existingRows[0].dataset.idx, ocrItems[0]);
    usedFirstRow = true;
  }

  // เพิ่มรายการสินค้าที่เหลือ ทีละแถว โดยใช้ createOrderItemRow/wireItemRow เดิม — เติมครบทุกช่อง
  ocrItems.slice(usedFirstRow ? 1 : 0).forEach(it => {
    const { row, idx } = createOrderItemRow();
    itemsList.appendChild(row);
    wireItemRowGlobal(idx);
    fillItemFields(idx, it);
  });

  if (itemsList.querySelectorAll('.order-item').length === 0) addItemRowGlobal();

  toast(`เติมข้อมูลจาก AI แล้ว ${ocrItems.length} รายการ — กรุณาเลือกหมวดสินค้าและตรวจสอบตัวเลขก่อนกดบันทึก`);
  if (typeof updateAllocPreview === 'function') updateAllocPreview();
}

function fillItemFields(idx, it){
   // ใช้หมวดสินค้าที่ผู้ใช้เลือก/พิมพ์มาจากหน้า OCR ถ้ามี ไม่งั้นเว้นว่างให้เลือกเองในฟอร์ม
  document.getElementById(`fCategory-${idx}`).value = it.category_name || '';
  document.getElementById(`fGroup-${idx}`).value = it.product_name || '';
  document.getElementById(`fProductUrl-${idx}`).value = it.product_url || '';
  document.getElementById(`fQty-${idx}`).value = it.qty || 1;
  document.getElementById(`fUnitBuy-${idx}`).value = it.unit_buy || 'ชิ้น';
  document.getElementById(`fRatio-${idx}`).value = it.ratio || 1;
  document.getElementById(`fUnitSell-${idx}`).value = it.unit_sell || 'ชิ้น';
  if (it.unit_price) document.getElementById(`fUnitPrice-${idx}`).value = it.unit_price;
  document.getElementById(`fItemPrice-${idx}`).value = it.total_price || 0;
}
</script>

## End

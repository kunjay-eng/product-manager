// ============================================================
// webapp_00_main.gs — Stock Manager Web App: Entry Point + บันทึก
// ============================================================
// 📁 โครงสร้างไฟล์ทั้งหมดของ "เว็บแอป" (แยกจากไฟล์ Telegram Bot เดิม
//    เช่น config.gs/data.gs/atr_report.gs/update.gs ที่ยังใช้ร่วมกันอยู่
//    — ไฟล์กลุ่ม webapp_* นี้เป็นแค่ "หน้าเว็บ" ที่เรียกใช้ฟังก์ชันเดิม
//    ไม่ duplicate logic ใดๆ):
//
//   webapp_00_main.gs      → จุดนี้ (doGet + ฟอร์มบันทึก: หุ้น US/TH/ปันผล)
//   webapp_01_summary.gs   → หน้า Summary (มูลค่าพอร์ตรวม, breakdown, alerts)
//   webapp_02_holdings.gs  → หน้า Holdings (ยอดถือครอง, อัปเดตราคา, Financial Goal)
//   webapp_03_analyze.gs   → หน้าวิเคราะห์ (Portfolio-mode deep dive + Fast signal)
//   webapp_04_dividend.gs  → หน้าปันผล (รายงานปันผลรายปี, เทรนด์รายเดือน)
//   webapp_05_settings.gs  → เฟส 1: Settings + StockMode (mode/group/risk ต่อหุ้น)
//   webapp_06_sparkline.gs → กราฟเส้นเล็กในการ์ด Holdings (จากไฟล์ Daily_Close_Log)
//
//   ไฟล์เดิมที่ยังใช้อยู่ (ไม่ได้แก้ในรอบนี้): config.gs, data.gs, atr_report.gs,
//   update.gs, summary_dashboard.gs, dashboard.gs, fx.gs, stockinfo.gs ฯลฯ
// ============================================================

// ──────────────────────────────────────
// Entry Point ของ Web App
// (doPost ของ Telegram Bot อยู่ใน bot.gs แยกกัน ไม่ชนกัน)
// ──────────────────────────────────────
//  function doGet() {
//  return HtmlService.createHtmlOutputFromFile('Index')
//    .setTitle('Stock Trade Entry')
//    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
//}


//function doGet() {
//  return HtmlService.createTemplateFromFile('Index')
//    .evaluate()
//    .setTitle('Stock Portfolio Manager')
//    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
//    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
//}

//**ข้อจำกัดที่ต้องรู้:** `Session.getActiveUser().getEmail()` จะได้ค่าว่างเปล่าถ้า "Who has access" ตั้งเป็น "Anyone" (Google ไม่ให้ identify ผู้ใช้ในโหมดนั้น) — โค้ดนี้จึงเป็นแค่**เกราะชั้นสอง** ไม่ใช่ตัวหลัก ต้องตั้ง "Only myself" ที่ Deployment ก่อนเป็นอันดับแรกเสมอ เพิ่มเช็คซ้ำในโค้ด (กันเผื่อตั้ง Deployment พลาด)

//  ## ข้อ 5 — สิทธิ์การเข้าถึง
//
//  การ Apps Script Web App จำกัดสิทธิ์หลักๆ ทำที่ **Deployment settings** (ไม่ใช่ในโค้ด) เช็คตามนี้ก่อน:
//
//  1. เปิด Apps Script Editor → **Deploy** → **Manage deployments** → คลิกดินสอ (แก้ไข) deployment ที่ใช้งานจริง
//  2. ดู 2 ช่องนี้:
//   - **Execute as:** ควรเป็น **"Me"** (คุณเอง) — ไม่ว่าใครเปิดลิงก์ สคริปต์จะรันด้วยสิทธิ์ของคุณเสมอ (เข้าถึง Sheet ได้)
//   - **Who has access:** ควรเป็น **"Only myself"** — ถ้าเป็น "Anyone" หรือ "Anyone with Google account" = ใครก็เปิดดูข้อมูลการเงินคุณได้ทันทีถ้ามีลิงก์

//ถ้าปัจจุบันตั้งเป็น "Anyone" อยู่ **แนะนำเปลี่ยนเป็น "Only myself" ทันที** แล้วกด Deploy ใหม่ (New version)


function doGet(e) {
  const ALLOWED_EMAIL = 'kunjay@gmail.com'; // ← ใส่อีเมล Google ของคุณ

  const currentUser = Session.getActiveUser().getEmail();
  if (currentUser && currentUser !== ALLOWED_EMAIL) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif; padding:40px; text-align:center">' +
      '<h2>🚫 ไม่มีสิทธิ์เข้าถึง</h2>' +
      '<p>บัญชีนี้ไม่ได้รับอนุญาตให้ใช้งานแอปนี้</p></div>'
    );
  }
  return HtmlService.createTemplateFromFile('Index')
   .evaluate()
   .setTitle('Stock Portfolio Manager')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ──────────────────────────────────────
// หา row ว่างถัดไป โดยเช็คจาก col ที่ระบุ
// ──────────────────────────────────────
function _getNextEmptyRow(sheet, checkCol, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;

  const colVals = sheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
  for (let i = 0; i < colVals.length; i++) {
    if (!colVals[i][0]) return startRow + i;
  }
  return lastRow + 1;
}

// ──────────────────────────────────────
// บันทึกหุ้นสหรัฐ — 🇺🇸 หุ้นสหรัฐ (USD)
// คอลัมน์: B=วันที่ C=หุ้น D=ประเภท E=จำนวนหุ้น F=ราคา/หุ้น
//          H=คอมมิชชัน I=ค่าธรรมเนียม J=TAF Fee K=VAT
// ──────────────────────────────────────
// ============================================================
// แทนที่ saveUSStock() และ saveTHStock() เดิมใน webapp_00_main.gs
// ทั้งฟังก์ชัน ด้วยเวอร์ชันนี้ (เหมือนเดิมทุกบรรทัด + เพิ่ม hook เรียก
// checkAndProcessSellCycle / processBuyAndCheckStockMode ก่อน return)
// ============================================================

function saveUSStock(data) {
  try {
    const sheet = getSheet(SHEETS.US_TRANS);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SHEETS.US_TRANS);

    const row = _getNextEmptyRow(sheet, 3, START_ROW.HOLD);
    const dateVal = new Date(data.date);

    sheet.getRange(row, 2).setValue(dateVal).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 3).setValue(data.ticker.toUpperCase());
    sheet.getRange(row, 4).setValue(data.type);
    sheet.getRange(row, 5).setValue(parseFloat(data.shares));
    sheet.getRange(row, 6).setValue(parseFloat(data.price));
    sheet.getRange(row, 8).setValue(parseFloat(data.commission) || 0);
    sheet.getRange(row, 9).setValue(parseFloat(data.taf) || 0);
    sheet.getRange(row, 10).setValue(parseFloat(data.tafFee) || 0);
    sheet.getRange(row, 11).setValue(parseFloat(data.vat) || 0);
    sheet.getRange(row, 17).setValue(parseFloat(data.exchangeRate) || 0);

    sheet.getRange(row, 5).setNumberFormat('0.0000000');
    sheet.getRange(row, 6).setNumberFormat('"$"#,##0.000');
    sheet.getRange(row,  1, 6).setNumberFormat('#,##0.0000');

    // ── HOOK เฟส 1/2 ──
    let needStockMode = false;
    try {
      if (data.type === 'ซื้อ') {
        const smCheck = processBuyAndCheckStockMode(data.ticker, 'us');
        needStockMode = smCheck.success && smCheck.needStockMode;
      
          if (needStockMode) {
          try { updateSingleUSPrice(data.ticker); } catch (priceErr) { logError('saveUSStock_priceUpdate', priceErr); }
        }

      } else if (data.type === 'ขาย') {
        checkAndProcessSellCycle(data.ticker, 'us');
      }
    } catch (hookErr) {
      // ไม่ปล่อยให้ hook พังแล้วทำให้บันทึกซื้อ/ขายหลักเสียหายไปด้วย
      logError('saveUSStock_hook', hookErr);
    }

      _clearTransactionCache();   // ← ย้ายมาตรงนี้ วิ่งทุกครั้งไม่ว่า hook จะพังหรือไม่
    // ✅ เคลียร์ _sheetDataCache (helper.gs) + _holdingsDataMemo (webapp_02_holdings.gs)
    //    ด้วย — กันจุดอื่นในรอบ execution เดียวกันที่เรียก getBuyPlanForTicker()/
    //    getHoldingsData() ต่อจากฟังก์ชันนี้แล้วได้ข้อมูลเก่าค้างจาก cache
    _clearSheetDataCache(SHEETS.US_TRANS);
    _clearHoldingsDataMemo();
    return { success: true, row: row, needStockMode: needStockMode };

  } catch (e) {
    logError('saveUSStock', e);
    return { success: false, error: e.message };
  }
}


function saveTHStock(data) {
  try {
    const sheet = getSheet(SHEETS.TH_TRANS);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SHEETS.TH_TRANS);

    const row = _getNextEmptyRow(sheet, 3, START_ROW.HOLD);
    const dateVal = new Date(data.date);

    sheet.getRange(row, 2).setValue(dateVal).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 3).setValue(data.ticker.toUpperCase());
    sheet.getRange(row, 4).setValue(data.type);
    sheet.getRange(row, 5).setValue(parseFloat(data.shares));
    sheet.getRange(row, 6).setValue(parseFloat(data.price));
    sheet.getRange(row, 8).setValue(parseFloat(data.commission) || 0);
    sheet.getRange(row, 9).setValue(parseFloat(data.fee) || 0);
    sheet.getRange(row, 10).setValue(parseFloat(data.vat) || 0);

    sheet.getRange(row, 5).setNumberFormat('#,##0');
    sheet.getRange(row, 6).setNumberFormat('#,##0.00');
    sheet.getRange(row,  1, 5).setNumberFormat('#,##0.00');

    // ── HOOK เฟส 1/2 ──
    let needStockMode = false;
    try {
      if (data.type === 'ซื้อ') {
        const smCheck = processBuyAndCheckStockMode(data.ticker, 'th');
        needStockMode = smCheck.success && smCheck.needStockMode;
            
      if (needStockMode) {
      try { updateSingleThaiPrice(data.ticker); } catch (priceErr) { logError('saveTHStock_priceUpdate', priceErr); }
        }

      } else if (data.type === 'ขาย') {
        checkAndProcessSellCycle(data.ticker, 'th');
      }
    } catch (hookErr) {
      logError('saveTHStock_hook', hookErr);
      
}
     _clearTransactionCache();  // ← ย้ายมาตรงนี้ วิ่งทุกครั้งไม่ว่า hook จะพังหรือไม่
    // ✅ เคลียร์ _sheetDataCache + _holdingsDataMemo เหมือน saveUSStock()
    _clearSheetDataCache(SHEETS.TH_TRANS);
    _clearHoldingsDataMemo();
    return { success: true, row: row, needStockMode: needStockMode };

  } catch (e) {
    logError('saveTHStock', e);
    return { success: false, error: e.message };
  }
}

// ──────────────────────────────────────
// บันทึกปันผล — 📝 บันทึกปันผล
// คอลัมน์: C=วันที่ XD D=วันที่รับปันผล E=Ticker J=ปันผล/หุ้น K=ปันผลรับรวม
// (ใช้ START_ROW.DIV จาก config.gs เดิม — ไม่ตั้งค่าใหม่)
// ──────────────────────────────────────
function saveDividend(data) {
  try {
    const sheet = getSheet(SHEETS.DIV);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SHEETS.DIV);

    const row = _getNextEmptyRow(sheet, 3, START_ROW.DIV);
    const xdDate  = new Date(data.xdDate);
    const recDate = new Date(data.recDate);

    sheet.getRange(row, 3).setValue(xdDate).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 4).setValue(recDate).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 5).setValue(data.ticker.toUpperCase());
    sheet.getRange(row, 10).setValue(parseFloat(data.divPerShare) || 0);
    sheet.getRange(row, 11).setValue(parseFloat(data.divTotal) || 0);

    sheet.getRange(row, 10).setNumberFormat('#,##0.00000');
    sheet.getRange(row, 11).setNumberFormat('#,##0.00');

    // ── เขียนภาษีหัก ณ ที่จ่าย (USD) ที่กรอกเอง แทนสูตร Excel เดิม ──
    const taxUsd = parseFloat(data.taxWithheld) || 0;
    sheet.getRange(row, DIV_COL.TAX_USD).setValue(taxUsd).setNumberFormat('#,##0.00');

    return { success: true, row: row };
  } catch (e) {
    logError('saveDividend', e);
    return { success: false, error: e.message };
  }
}



// ============================================================
// วางไฟล์นี้ (หรือแปะท้าย webapp_00_main.gs ก็ได้) แล้วเลือกฟังก์ชัน
// testSaveUSStock / testSaveTHStock จาก dropdown ข้าง ▶ เรียกใช้ แทนที่
// จะเรียก saveUSStock ตรงๆ — ฟังก์ชันนี้จะส่ง data ปลอมเข้าไปให้ครบ
// ทำให้ debug ได้จริงว่า error เกิดตรงไหน (รวมถึง hook เฟส 1/2 ด้วย)
// ============================================================

function testSaveUSStock() {
  const result = saveUSStock({
    date: '2026-07-15',
    ticker: 'WDC',   // ⚠️ ใช้ ticker ที่ไม่มีจริงกันพลาดข้อมูลจริง
    type: 'ซื้อ',        // ลองเปลี่ยนเป็น 'ขาย' เพื่อเทส hook เฟส 1 (cycle)
    shares: '0.1089637',
    price: '100',
    commission: '0',
    taf: '0',
    tafFee: '0',
    vat: '0',
    exchangeRate: '33.5'
  });
  Logger.log(JSON.stringify(result));
}

function testSaveTHStock() {
  const result = saveTHStock({
    date: '2026-07-15',
    ticker: 'TESTTH',
    type: 'ซื้อ',
    shares: '100',
    price: '10',
    commission: '0',
    fee: '0',
    vat: '0'
  });
  Logger.log(JSON.stringify(result));
}



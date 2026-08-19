// ============================================================
// webapp_10_stock_lifecycle.gs — รวมเฟส 1/2/3 เป็นไฟล์เดียว
// (แทนที่ webapp_10_sell_cycle.gs + webapp_11_buy_stockmode.gs +
//  webapp_12_cleanup.gs ทั้ง 3 ไฟล์เดิม — ลบ 3 ไฟล์นั้นออกก่อนวางไฟล์นี้
//  เพื่อกัน const ประกาศซ้ำ)
//
// อ้างอิง constants จาก config.gs จริงทั้งหมด: SHEETS, START_ROW, HOLD_COL,
// ATR_SHEETS, ATR_START_ROW, BACKEND_ATR — ไม่ประกาศชื่อชีต/คอลัมน์ใหม่ซ้ำ
//
// ⚠️ แก้ไข: BACKEND_ATR.TH_TICKER/TH_ATR ใน config.gs ปัจจุบันตั้งเป็น 3,4
// (col C,D) แต่จากภาพชีตจริง Backend_ATR ฝั่งไทยอยู่ col E,F (5,6) — ไฟล์นี้
// ใช้ 5,6 ตรงๆ (ไม่พึ่ง BACKEND_ATR.TH_TICKER/TH_ATR) แนะนำให้ไปแก้ค่าใน
// config.gs ให้ตรงกันด้วยเพื่อกันงงในอนาคต
//
// ⚠️ StockMode: config.gs ที่ส่งมาไม่มี object สำหรับชีตนี้ ไฟล์นี้เลย
// ประกาศเองด้านล่าง (SM_SHEET_NAME, SM_START_ROW, SM_COL) — ถ้ามีอยู่แล้ว
// ในไฟล์อื่นของโปรเจกต์ให้ลบของไฟล์นี้ทิ้งแล้วใช้ของเดิมแทน
// ============================================================

const SM_SHEET_NAME = 'StockMode';
const SM_START_ROW   = 7;   // header แถว 6 → ข้อมูลเริ่มแถว 7
const SM_COL = { TICKER: 2, MARKET: 3, ASSET_TYPE: 4, MODE: 5, TREND_GROUP: 6,
                  TARGET_WEIGHT_OVERRIDE: 8, CUT_STOP: 9, TRAILING_STOP: 10,
                  TAKE_PROFIT: 11, NOTE: 12,
                  PORTFOLIO_TRAIL_START: 13, PORTFOLIO_TRAIL_ATR_X: 14,
                  MIN_PROFIT_PROTECT: 15, TRAIL_TRIGGER_MODE: 16, TRAIL_RESET_PCT: 17 };



// Realized P&L: config.gs ไม่มี object คอลัมน์แยก แต่ ticker อยู่ col B (2)
// เหมือน HOLD_COL.TICKER (โครงสร้างเดียวกับ Holdings)
const REAL_COL_TICKER = 2;

// ============================================================
// เฟส 1 — บันทึกขายหุ้น + ระบบ Cycle (_c1, _c2, ...)
// ============================================================

/**
 * เรียกทันทีหลังบันทึกขายหุ้นสำเร็จ (จาก saveUSStock/saveTHStock เมื่อ
 * data.type === 'ขาย')
 */
function checkAndProcessSellCycle(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'ไม่ระบุ ticker' };

    const holdSheetName = market === 'th' ? SHEETS.TH_HOLD : SHEETS.US_HOLD;
    const holdSheet = getSheet(holdSheetName);
    if (!holdSheet) throw new Error('ไม่พบ sheet: ' + holdSheetName);

    const lastRow = holdSheet.getLastRow();
    if (lastRow < START_ROW.HOLD) return { success: true, cycled: false };

    const tickers = holdSheet
      .getRange(START_ROW.HOLD, HOLD_COL.TICKER, lastRow - START_ROW.HOLD + 1, 1)
      .getValues();

    let holdRow = -1;
    for (let i = 0; i < tickers.length; i++) {
      if (String(tickers[i][0]).trim().toUpperCase() === ticker) { holdRow = START_ROW.HOLD + i; break; }
    }
    if (holdRow === -1) return { success: true, cycled: false };

    const remain = parseFloat(holdSheet.getRange(holdRow, HOLD_COL.SHARES_REMAIN).getValue()) || 0;
    if (Math.abs(remain) > 0.0000001) return { success: true, cycled: false };

    const newTicker = getNextCycleTicker(ticker, market);
    renameTickerAcrossSheets(ticker, newTicker, market);
    clearStockModeRow(ticker);

    return { success: true, cycled: true, oldTicker: ticker, newTicker: newTicker };
  } catch (e) {
    logError('checkAndProcessSellCycle', e);
    return { success: false, error: e.message };
  }
}

function getNextCycleTicker(baseTicker, market) {
  const transSheetName = market === 'th' ? SHEETS.TH_TRANS : SHEETS.US_TRANS;
  const sheet = getSheet(transSheetName);
  const lastRow = sheet.getLastRow();
  const startRow = START_ROW.HOLD;
  if (lastRow < startRow) return baseTicker + '_c1';

  const values = sheet.getRange(startRow, 3, lastRow - startRow + 1, 1).getValues(); // col C = ticker

  const pattern = new RegExp('^' + baseTicker + '_C(\\d+)$');
  let maxCycle = 0;
  values.forEach(r => {
    const v = String(r[0] || '').trim().toUpperCase();
    const m = v.match(pattern);
    if (m) maxCycle = Math.max(maxCycle, parseInt(m[1], 10));
  });
  return baseTicker + '_c' + (maxCycle + 1);
}

function renameTickerAcrossSheets(oldTicker, newTicker, market) {
  const transSheetName = market === 'th' ? SHEETS.TH_TRANS : SHEETS.US_TRANS;
  const holdSheetName  = market === 'th' ? SHEETS.TH_HOLD  : SHEETS.US_HOLD;
  const realSheetName  = market === 'th' ? SHEETS.TH_REAL  : SHEETS.US_REAL;

  _renameInSheet(getSheet(transSheetName), 3, START_ROW.HOLD, oldTicker, newTicker);
  _renameInSheet(getSheet(holdSheetName),  HOLD_COL.TICKER, START_ROW.HOLD, oldTicker, newTicker);
  _renameInSheet(getSheet(realSheetName),  REAL_COL_TICKER, START_ROW.REALIZED, oldTicker, newTicker);
}

function _renameInSheet(sheet, colIdx, startRow, oldTicker, newTicker) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, colIdx, lastRow - startRow + 1, 1);
  const values = range.getValues();
  let changed = false;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === oldTicker) {
      values[i][0] = newTicker;
      changed = true;
    }
  }
  if (changed) range.setValues(values);
}

function clearStockModeRow(ticker) {
  const sheet = getSheet(SM_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < SM_START_ROW) return;

  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(SM_START_ROW, SM_COL.TICKER, lastRow - SM_START_ROW + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === ticker) {
      sheet.getRange(SM_START_ROW + i, 1, 1, lastCol).clearContent();
      break;
    }
  }
}

// ============================================================
// เฟส 2 — บันทึกซื้อหุ้น + StockMode Step 2
// ============================================================

function processBuyAndCheckStockMode(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'ไม่ระบุ ticker' };

    const holdSheetName = market === 'th' ? SHEETS.TH_HOLD : SHEETS.US_HOLD;
    const realSheetName = market === 'th' ? SHEETS.TH_REAL : SHEETS.US_REAL;

    _ensureTickerRow(getSheet(holdSheetName), HOLD_COL.TICKER, START_ROW.HOLD, ticker);
    _ensureTickerRow(getSheet(realSheetName), REAL_COL_TICKER, START_ROW.REALIZED, ticker);

    const needStockMode = !_tickerExistsInStockMode(ticker);
    return { success: true, needStockMode: needStockMode, ticker: ticker };
  } catch (e) {
    logError('processBuyAndCheckStockMode', e);
    return { success: false, error: e.message };
  }
}

function _ensureTickerRow(sheet, colIdx, startRow, ticker) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();

  if (lastRow >= startRow) {
    const values = sheet.getRange(startRow, colIdx, lastRow - startRow + 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim().toUpperCase() === ticker) return;
    }
  }
  const targetRow = _getNextEmptyRow(sheet, colIdx, startRow);
  sheet.getRange(targetRow, colIdx).setValue(ticker);
}

function _tickerExistsInStockMode(ticker) {
  const sheet = getSheet(SM_SHEET_NAME);
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < SM_START_ROW) return false;

  const values = sheet.getRange(SM_START_ROW, SM_COL.TICKER, lastRow - SM_START_ROW + 1, 1).getValues();
  return values.some(r => String(r[0]).trim().toUpperCase() === ticker);
}

function saveNewStockMode(data) {
  try {
    const ticker = String(data.ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'ไม่ระบุ ticker' };
    if (!data.mode) return { success: false, error: 'กรุณาเลือกโหมด (Fast/Portfolio)' };
    if (!data.trendGroup) return { success: false, error: 'กรุณาเลือกกลุ่ม Trend' };

    const sheet = getSheet(SM_SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SM_SHEET_NAME);
    if (_tickerExistsInStockMode(ticker)) return { success: false, error: ticker + ' มีอยู่ใน StockMode แล้ว' };

    const row = _getNextEmptyRow(sheet, SM_COL.TICKER, SM_START_ROW);
    const marketLabel = data.market === 'th' ? 'ไทย' : 'สหรัฐ';

    sheet.getRange(row, SM_COL.TICKER).setValue(ticker);
    sheet.getRange(row, SM_COL.MARKET).setValue(marketLabel);
    sheet.getRange(row, SM_COL.ASSET_TYPE).setValue(data.assetType || 'หุ้น');
    sheet.getRange(row, SM_COL.MODE).setValue(data.mode);
    sheet.getRange(row, SM_COL.TREND_GROUP).setValue(data.trendGroup);

   if (data.targetWeightOverride !== '' && data.targetWeightOverride != null)
  sheet.getRange(row, SM_COL.TARGET_WEIGHT_OVERRIDE).setValue(parseFloat(data.targetWeightOverride));
if (data.cutStopOverride !== '' && data.cutStopOverride != null)
  sheet.getRange(row, SM_COL.CUT_STOP).setValue(parseFloat(data.cutStopOverride));
if (data.trailingStopOverride !== '' && data.trailingStopOverride != null)
  sheet.getRange(row, SM_COL.TRAILING_STOP).setValue(parseFloat(data.trailingStopOverride));
if (data.takeProfitOverride !== '' && data.takeProfitOverride != null)
  sheet.getRange(row, SM_COL.TAKE_PROFIT).setValue(parseFloat(data.takeProfitOverride) / 100);
if (data.portfolioTrailStartOverride !== '' && data.portfolioTrailStartOverride != null)
  sheet.getRange(row, SM_COL.PORTFOLIO_TRAIL_START).setValue(parseFloat(data.portfolioTrailStartOverride));
if (data.portfolioTrailAtrOverride !== '' && data.portfolioTrailAtrOverride != null)
  sheet.getRange(row, SM_COL.PORTFOLIO_TRAIL_ATR_X).setValue(parseFloat(data.portfolioTrailAtrOverride));

if (data.note) sheet.getRange(row, SM_COL.NOTE).setValue(data.note);


    return { success: true, row: row };
  } catch (e) {
    logError('saveNewStockMode', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// แก้ไขหุ้นที่มีอยู่แล้วในชีต StockMode (จากหน้าเว็บ #page-settings)
// ══════════════════════════════════════════════════════════
function updateStockMode(data) {
  try {
    const ticker = String(data.ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'ไม่ระบุ ticker' };

    const sheet = getSheet(SM_SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SM_SHEET_NAME);

    const lastRow = sheet.getLastRow();
    if (lastRow < SM_START_ROW) return { success: false, error: 'ไม่พบ ' + ticker + ' ใน StockMode' };

    const numRows = lastRow - SM_START_ROW + 1;
    const tickerCol = sheet.getRange(SM_START_ROW, SM_COL.TICKER, numRows, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < tickerCol.length; i++) {
      if (String(tickerCol[i][0] || '').trim().toUpperCase() === ticker) { targetRow = SM_START_ROW + i; break; }
    }
    if (targetRow === -1) return { success: false, error: 'ไม่พบ ' + ticker + ' ใน StockMode' };

    sheet.getRange(targetRow, SM_COL.MARKET).setValue(data.market || 'สหรัฐ');
    sheet.getRange(targetRow, SM_COL.ASSET_TYPE).setValue(data.assetType || 'หุ้น');
    sheet.getRange(targetRow, SM_COL.MODE).setValue(data.mode || 'Portfolio');
    sheet.getRange(targetRow, SM_COL.TREND_GROUP).setValue(data.trendGroup || '');

    sheet.getRange(targetRow, SM_COL.TARGET_WEIGHT_OVERRIDE)
      .setValue(data.targetWeightOverride !== '' && data.targetWeightOverride != null ? parseFloat(data.targetWeightOverride) : '');
    sheet.getRange(targetRow, SM_COL.CUT_STOP)
      .setValue(data.cutStopOverride !== '' && data.cutStopOverride != null ? parseFloat(data.cutStopOverride) : '');
    sheet.getRange(targetRow, SM_COL.TRAILING_STOP)
      .setValue(data.trailingStopOverride !== '' && data.trailingStopOverride != null ? parseFloat(data.trailingStopOverride) : '');
    sheet.getRange(targetRow, SM_COL.TAKE_PROFIT)
      .setValue(data.takeProfitOverride !== '' && data.takeProfitOverride != null ? parseFloat(data.takeProfitOverride) / 100 : '');
    sheet.getRange(targetRow, SM_COL.NOTE).setValue(data.note || '');

 sheet.getRange(targetRow, SM_COL.PORTFOLIO_TRAIL_START)
  .setValue(data.portfolioTrailStartOverride !== '' && data.portfolioTrailStartOverride != null ? parseFloat(data.portfolioTrailStartOverride) : '');
sheet.getRange(targetRow, SM_COL.PORTFOLIO_TRAIL_ATR_X)
  .setValue(data.portfolioTrailAtrOverride !== '' && data.portfolioTrailAtrOverride != null ? parseFloat(data.portfolioTrailAtrOverride) : '');
sheet.getRange(targetRow, SM_COL.MIN_PROFIT_PROTECT)
  .setValue(data.minProfitProtectOverride !== '' && data.minProfitProtectOverride != null ? parseFloat(data.minProfitProtectOverride) : '');
sheet.getRange(targetRow, SM_COL.TRAIL_TRIGGER_MODE)
  .setValue(data.trailTriggerModeOverride || '');
sheet.getRange(targetRow, SM_COL.TRAIL_RESET_PCT)
  .setValue(data.trailResetPctOverride !== '' && data.trailResetPctOverride != null ? parseFloat(data.trailResetPctOverride) : '');



    return { success: true, row: targetRow };
  } catch (e) {
    logError('updateStockMode', e);
    return { success: false, error: e.message };
  }
}



// ============================================================
// เฟส 3 — ปุ่ม Cleanup (Settings)
// ============================================================

function runCleanup() {
  try {
    const backendCount = _cleanupBackendATR();
    const usCount = _cleanupATRPortfolio(ATR_SHEETS.US, ATR_START_ROW.US);
    const thCount = _cleanupATRPortfolio(ATR_SHEETS.TH, ATR_START_ROW.TH);

    return {
      success: true,
      backendCount: backendCount,
      usCount: usCount,
      thCount: thCount,
      total: backendCount + usCount + thCount
    };
  } catch (e) {
    logError('runCleanup', e);
    return { success: false, error: e.message };
  }
}

/**
 * Backend_ATR: col A ว่าง → ล้าง col B (ฝั่งสหรัฐ)
 *              col E ว่าง → ล้าง col F (ฝั่งไทย) ⚠️ ยืนยันจากภาพจริงแล้วว่าเป็น E/F
 *              (ไม่ใช้ BACKEND_ATR.TH_TICKER/TH_ATR จาก config.gs เพราะค่านั้นผิด)
 */
function _cleanupBackendATR() {
  const sheet = getSheet(BACKEND_ATR.SHEET);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < BACKEND_ATR.START_ROW) return 0;

  const numRows = lastRow - BACKEND_ATR.START_ROW + 1;
  let cleared = 0;

  // ── ฝั่งสหรัฐ: ใช้ BACKEND_ATR.US_TICKER/US_ATR จาก config (ถูกอยู่แล้ว: A/B) ──
  const colA = sheet.getRange(BACKEND_ATR.START_ROW, BACKEND_ATR.US_TICKER, numRows, 1).getValues();
  const colB = sheet.getRange(BACKEND_ATR.START_ROW, BACKEND_ATR.US_ATR, numRows, 1).getValues();
  for (let i = 0; i < numRows; i++) {
    if (!String(colA[i][0]).trim() && String(colB[i][0]).trim() !== '') {
      sheet.getRange(BACKEND_ATR.START_ROW + i, BACKEND_ATR.US_ATR).clearContent();
      cleared++;
    }
  }

  // ── ฝั่งไทย: E(5) ว่าง → ล้าง F(6) — ยืนยันจากภาพจริง ──
  const colE = sheet.getRange(BACKEND_ATR.START_ROW, 5, numRows, 1).getValues();
  const colF = sheet.getRange(BACKEND_ATR.START_ROW, 6, numRows, 1).getValues();
  for (let i = 0; i < numRows; i++) {
    if (!String(colE[i][0]).trim() && String(colF[i][0]).trim() !== '') {
      sheet.getRange(BACKEND_ATR.START_ROW + i, 6).clearContent();
      cleared++;
    }
  }

  return cleared;
}

/**
 * ATR_Portfolio (US/TH): col A (Symbol) ว่าง → ล้าง col B (โหมด) และ col D (Highest Close)
 */
function _cleanupATRPortfolio(sheetName, startRow) {
  const sheet = getSheet(sheetName);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return 0;

  const numRows = lastRow - startRow + 1;
  const colA = sheet.getRange(startRow, 1, numRows, 1).getValues();
  const colB = sheet.getRange(startRow, 2, numRows, 1).getValues();
  const colD = sheet.getRange(startRow, 4, numRows, 1).getValues();

  let cleared = 0;
  for (let i = 0; i < numRows; i++) {
    const aEmpty  = !String(colA[i][0]).trim();
    const bHasVal = String(colB[i][0]).trim() !== '';
    const dHasVal = String(colD[i][0]).trim() !== '';
    if (aEmpty && (bHasVal || dHasVal)) {
      const row = startRow + i;
      sheet.getRange(row, 2).clearContent();
      sheet.getRange(row, 4).clearContent();
      cleared++;
    }
  }
  return cleared;
}

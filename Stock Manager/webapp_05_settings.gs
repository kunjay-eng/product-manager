// ============================================================
// webapp_05_settings.gs — เฟส 1: Settings + StockMode
// โครงสร้างพื้นฐาน mode/group/risk รายตัวหุ้น สำหรับ webapp_03_analyze.gs
// (Fast-mode) และ webapp_01_summary.gs (rebalance alert)
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
//
// วิธีติดตั้ง (ทำครั้งเดียว ถ้ายังไม่เคยรัน):
//   เปิดไฟล์นี้ เลือกฟังก์ชัน "setupNewSheets" จากดรอปดาวน์ด้านบน
//   (ข้าง Debug) แล้วกด ▶ Run เพื่อสร้างชีต Settings + StockMode
// ============================================================





const SETTINGS_SHEET = {
  SHEET:               'Settings',
  CUT_STOP:            'E4',  // Cut Stop % (ค่าติดลบ)
  TRAIL_START_PROFIT:  'E5',  // เริ่ม Trailing Stop เมื่อกำไรเกิน %
  TRAILING_STOP:       'E6',  // Trailing Stop %
  TAKE_PROFIT_MIN:     'E7',  // Take Profit เตือน ขั้นต่ำ %
  TAKE_PROFIT_MAX:     'E8',  // Take Profit เตือน สูงสุด %
  CONCENTRATION_WARN:  'E9',  // เตือน Concentration Risk ถ้า > % นี้ต่อตัว/กลุ่ม
  DRAWDOWN_WARN:       'E10'  // เตือน Drawdown จากจุดสูงสุด
};

// ── คอลัมน์ชีต StockMode (1-based) — 1 แถวต่อ 1 ticker ──
const STOCK_MODE_COL = {
  TICKER: 2, MARKET: 3, ASSET_TYPE: 4, MODE: 5, TREND_GROUP: 6,
  TARGET_WEIGHT_AUTO: 7, TARGET_WEIGHT_OVERRIDE: 8,
  CUT_STOP_ATR_X: 9, TRAIL_STOP_ATR_X: 10, TAKE_PROFIT_PCT: 11, NOTE: 12,
  PORTFOLIO_TRAIL_START: 13, PORTFOLIO_TRAIL_ATR_X: 14,
  MIN_PROFIT_PROTECT: 15, TRAIL_TRIGGER_MODE: 16, TRAIL_RESET_PCT: 17
};

const STOCK_MODE_SHEET_NAME = 'StockMode';
const STOCK_MODE_START_ROW  = 7; // แถวแรกที่กรอกข้อมูลหุ้น (header อยู่แถว 6)

// ตัวเลือก dropdown — เก็บเป็น constant แทนชีต "option" แยกต่างหาก
// เพราะหน้าเว็บไม่จำเป็นต้องพึ่ง data validation ของ Google Sheets
// (แต่ยังใช้สร้าง dropdown ในชีต StockMode ให้กรอกง่ายขึ้นด้วย)
const STOCK_MODE_OPTIONS = {
  MARKET:      ['สหรัฐ', 'ไทย'],
  ASSET_TYPE:  ['หุ้น', 'ETF', 'กองทุน'],
  MODE:        ['Fast', 'Portfolio'],
  TREND_GROUP: ['Growth', 'Value', 'Defensive', 'Dividend']
};

// ══════════════════════════════════════════════════════════
// ติดตั้งครั้งเดียว — รันจาก Apps Script editor (ดูวิธีด้านบนของไฟล์)
// ══════════════════════════════════════════════════════════
function setupNewSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const createdSettings  = _setupSettingsSheet(ss);
  const createdStockMode = _setupStockModeSheet(ss);

  const msg = (createdSettings ? '✅ สร้างชีต Settings แล้ว\n' : 'ℹ️ ชีต Settings มีอยู่แล้ว ไม่ได้แก้ไข\n') +
              (createdStockMode ? '✅ สร้างชีต StockMode แล้ว' : 'ℹ️ ชีต StockMode มีอยู่แล้ว ไม่ได้แก้ไข');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function _setupSettingsSheet(ss) {
  if (ss.getSheetByName(SETTINGS_SHEET.SHEET)) return false; // มีอยู่แล้ว ไม่ทับ

  const sheet = ss.insertSheet(SETTINGS_SHEET.SHEET);
  sheet.getRange('B2')
    .setValue('⚙️ Settings — ค่า Risk Management เริ่มต้น (override รายตัวได้ที่ชีต StockMode)')
    .setFontWeight('bold').setFontSize(12);

  const rows = [
    ['B4',  'Cut Stop (%)',                                    'E4',  -0.07],
    ['B5',  'เริ่ม Trailing Stop เมื่อกำไรเกิน (%)',            'E5',   0.05],
    ['B6',  'Trailing Stop (%)',                                'E6',   0.08],
    ['B7',  'Take Profit เตือน — ขั้นต่ำ (%)',                  'E7',   0.15],
    ['B8',  'Take Profit เตือน — สูงสุด (%)',                   'E8',   0.20],
    ['B9',  'เตือน Concentration Risk ถ้าตัวใดตัวหนึ่ง > (%)',  'E9',   0.25],
    ['B10', 'เตือน Drawdown จากจุดสูงสุด (%)',                  'E10', -0.15]
  ];
  rows.forEach(([labelCell, label, valueCell, value]) => {
    sheet.getRange(labelCell).setValue(label);
    sheet.getRange(valueCell).setValue(value).setNumberFormat('0.00%');
  });

  sheet.setColumnWidth(2, 340);
  sheet.getRange('E4:E10').setBackground('#fff2cc'); // เหลือง = แก้ไขได้
  return true;
}

function _setupStockModeSheet(ss) {
  if (ss.getSheetByName(STOCK_MODE_SHEET_NAME)) return false; // มีอยู่แล้ว ไม่ทับ

  const sheet = ss.insertSheet(STOCK_MODE_SHEET_NAME);
  sheet.getRange('B2')
    .setValue('🏷️ StockMode — กำหนดโหมด/กลุ่ม/เป้าหมายรายตัวหุ้น')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange('B3')
    .setValue('💡 กรอกเซลล์สีเหลือง ต่อหุ้นแต่ละตัว (ถ้าไม่กรอก ระบบใช้ค่า default จาก Settings) — หุ้นเริ่ม row ' + STOCK_MODE_START_ROW);

  const headers = [
    'Ticker', 'ตลาด', 'ประเภทสินทรัพย์', 'โหมด (Fast/Portfolio)', 'กลุ่ม Trend',
    'Target Weight - Auto (คำนวณ)', 'Target Weight - Override (กรอกเอง)',
    'Cut Stop ATR x (override)', 'Trailing Stop ATR x (override)',
    'Take Profit % (override)', 'หมายเหตุ'
  ];
  const headerRow = STOCK_MODE_START_ROW - 1;
  sheet.getRange(headerRow, STOCK_MODE_COL.TICKER, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold').setBackground('#e8eaf6').setWrap(true);

  const numDataRows = 100;
  _setDropdown(sheet, STOCK_MODE_START_ROW, STOCK_MODE_COL.MARKET,      numDataRows, STOCK_MODE_OPTIONS.MARKET);
  _setDropdown(sheet, STOCK_MODE_START_ROW, STOCK_MODE_COL.ASSET_TYPE,  numDataRows, STOCK_MODE_OPTIONS.ASSET_TYPE);
  _setDropdown(sheet, STOCK_MODE_START_ROW, STOCK_MODE_COL.MODE,        numDataRows, STOCK_MODE_OPTIONS.MODE);
  _setDropdown(sheet, STOCK_MODE_START_ROW, STOCK_MODE_COL.TREND_GROUP, numDataRows, STOCK_MODE_OPTIONS.TREND_GROUP);

  sheet.getRange(STOCK_MODE_START_ROW, STOCK_MODE_COL.TARGET_WEIGHT_OVERRIDE, numDataRows, 4).setBackground('#fff2cc');
  sheet.setColumnWidths(2, 11, 130);
  return true;
}

function _setDropdown(sheet, startRow, col, numRows, options) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule);
}

// ══════════════════════════════════════════════════════════
// อ่านค่า Risk Management เริ่มต้นจากชีต Settings
// (ถ้ายังไม่ได้รัน setupNewSheets() จะคืนค่า default เดียวกันแบบ hardcode)
// ══════════════════════════════════════════════════════════
function getRiskSettings() {
  const fallback = {
    cutStopPct: -0.07, trailStartProfitPct: 0.05, trailingStopPct: 0.08,
    takeProfitMinPct: 0.15, takeProfitMaxPct: 0.20,
    concentrationWarnPct: 0.25, drawdownWarnPct: -0.15
  };
  try {
    const sheet = getSheet(SETTINGS_SHEET.SHEET);
    return {
      cutStopPct:           Number(sheet.getRange(SETTINGS_SHEET.CUT_STOP).getValue())           || fallback.cutStopPct,
      trailStartProfitPct:  Number(sheet.getRange(SETTINGS_SHEET.TRAIL_START_PROFIT).getValue())  || fallback.trailStartProfitPct,
      trailingStopPct:      Number(sheet.getRange(SETTINGS_SHEET.TRAILING_STOP).getValue())       || fallback.trailingStopPct,
      takeProfitMinPct:     Number(sheet.getRange(SETTINGS_SHEET.TAKE_PROFIT_MIN).getValue())      || fallback.takeProfitMinPct,
      takeProfitMaxPct:     Number(sheet.getRange(SETTINGS_SHEET.TAKE_PROFIT_MAX).getValue())      || fallback.takeProfitMaxPct,
      concentrationWarnPct: Number(sheet.getRange(SETTINGS_SHEET.CONCENTRATION_WARN).getValue())   || fallback.concentrationWarnPct,
      drawdownWarnPct:      Number(sheet.getRange(SETTINGS_SHEET.DRAWDOWN_WARN).getValue())        || fallback.drawdownWarnPct
    };
  } catch (e) {
    logError('getRiskSettings', e);
    return fallback;
  }
}

// ══════════════════════════════════════════════════════════
// อ่านค่า mode/group/override รายตัวจากชีต StockMode
// return: { "VOO": {market, assetType, mode, trendGroup, targetWeightAuto,
//           targetWeightOverride, cutStopAtrX, trailStopAtrX,
//           takeProfitPctOverride, note}, ... }
// ══════════════════════════════════════════════════════════
function getStockModeMap() {
  const map = {};
  try {
    const sheet = getSheet(STOCK_MODE_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < STOCK_MODE_START_ROW) return map;

    const numRows = lastRow - STOCK_MODE_START_ROW + 1;
   const rows = sheet.getRange(STOCK_MODE_START_ROW, STOCK_MODE_COL.TICKER, numRows, 16).getValues();

rows.forEach(row => {
  const ticker = String(row[0] || '').trim().toUpperCase();
  if (!ticker) return;
  map[ticker] = {
    market: row[1] || '', assetType: row[2] || '', mode: row[3] || 'Portfolio',
    trendGroup: row[4] || '',
    targetWeightAuto: row[5] === '' ? null : Number(row[5]),
    targetWeightOverride: row[6] === '' ? null : Number(row[6]),
    cutStopAtrX: row[7] === '' ? null : Number(row[7]),
    trailStopAtrX: row[8] === '' ? null : Number(row[8]),
    takeProfitPctOverride: row[9] === '' ? null : Number(row[9]),
    note: row[10] || '',
    portfolioTrailStartOverride: row[11] === '' ? null : Number(row[11]),
    portfolioTrailAtrXOverride:  row[12] === '' ? null : Number(row[12]),
    minProfitProtectOverride: row[13] === '' ? null : Number(row[13]),
    trailTriggerModeOverride: row[14] || null,
    trailResetPctOverride: row[15] === '' ? null : Number(row[15])
  };
});


  } catch (e) {
    logError('getStockModeMap', e);
  }
  return map;
}

// ══════════════════════════════════════════════════════════
// รวมค่า override รายตัว (StockMode) เข้ากับค่า default (Settings)
// ให้ค่าสุดท้ายที่แต่ละหุ้นควรใช้จริง — ฟังก์ชันนี้จะถูกเรียกใช้โดย
// Analysis_Fast / Analysis_Portfolio ในเฟสถัดไป (ยังไม่ได้ต่อการใช้งานตอนนี้)
// ══════════════════════════════════════════════════════════
function getEffectiveRiskParams(ticker) {
  const defaults = getRiskSettings();
  const modeMap  = getStockModeMap();
  const cfg = modeMap[String(ticker).trim().toUpperCase()] || {};

  return {
    ticker: ticker,
    mode: cfg.mode || 'Portfolio',
    trendGroup: cfg.trendGroup || '',
    targetWeight: (cfg.targetWeightOverride !== null && cfg.targetWeightOverride !== undefined)
      ? cfg.targetWeightOverride
      : cfg.targetWeightAuto,

    // Portfolio mode ใช้ % แบบตายตัวจาก Settings เสมอ (ยังไม่มี override % ต่อตัวใน StockMode)
    cutStopPct: defaults.cutStopPct,
    trailStartProfitPct: defaults.trailStartProfitPct,
    trailingStopPct: defaults.trailingStopPct,

    // Fast mode ใช้ตัวคูณ ATR — null หมายถึงยังไม่ override ให้ใช้ multiplier เดิมจาก ATR_Portfolio
    cutStopAtrX:   (cfg.cutStopAtrX   !== null && cfg.cutStopAtrX   !== undefined) ? cfg.cutStopAtrX   : null,
    trailStopAtrX: (cfg.trailStopAtrX !== null && cfg.trailStopAtrX !== undefined) ? cfg.trailStopAtrX : null,

    takeProfitPct: (cfg.takeProfitPctOverride !== null && cfg.takeProfitPctOverride !== undefined)
      ? cfg.takeProfitPctOverride
      : defaults.takeProfitMinPct,

    concentrationWarnPct: defaults.concentrationWarnPct,
    drawdownWarnPct: defaults.drawdownWarnPct,

portfolioTrailStartProfitPct: cfg.portfolioTrailStartOverride ?? 10,
portfolioTrailAtrX: cfg.portfolioTrailAtrXOverride ?? 3.5,
minProfitProtectPct: cfg.minProfitProtectOverride ?? 1,
trailTriggerMode: cfg.trailTriggerModeOverride ?? 'lastClose',
trailResetPct: cfg.trailResetPctOverride ?? 0.5

  };
}

// ══════════════════════════════════════════════════════════
// เรียกจากหน้าเว็บ (google.script.run) — สำหรับหน้า "เพิ่มเติม/ตั้งค่า"
// ในเฟสถัดไป ตอนนี้เป็น read-only ก่อน (ยังไม่มี UI แก้ไขจากหน้าเว็บ)
// ══════════════════════════════════════════════════════════
function getSettingsData() {
  try {
    return { success: true, risk: getRiskSettings(), stockModes: getStockModeMap() };
  } catch (e) {
    try { logError('getSettingsData', e); } catch (logErr) { /* กัน logError พังซ้ำ */ }
    return { success: false, error: (e && e.message) ? e.message : String(e) || 'Unknown error (no message)' };
  }
}







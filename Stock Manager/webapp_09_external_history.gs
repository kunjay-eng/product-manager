// ============================================================
// webapp_09_external_history.gs — ดึงราคาย้อนหลังจาก Daily_Close_Log
// (spreadsheet แยกไฟล์ รันบันทึกทุกวันตี 2 เฉพาะหุ้นที่ถืออยู่)
// ใช้คำนวณ EMA50 + MACD เพิ่มเติมให้ Technical Checklist ในหน้าวิเคราะห์รายตัว
// (webapp_03c_stock_detail.gs) โดยไม่ต้องยิง Yahoo Finance ซ้ำ
//
// ⚠️ ต้องแก้ EXTERNAL_LOG_SHEET_ID ด้านล่างเป็น ID จริงของไฟล์ Daily_Close_Log
//    ก่อนใช้งาน — เอาจาก URL ไฟล์นั้น: .../spreadsheets/d/[ID นี้]/edit
// ⚠️ ข้อมูลมีเฉพาะหุ้นที่ถืออยู่เท่านั้น — หุ้นใน Watchlist/เพิ่ง IPO จะไม่มีข้อมูล
//    ระบบจะแจ้งชัดเจนว่า "ไม่มีข้อมูลพอ" แทนการเดา ไม่ปั้นตัวเลขขึ้นมา
// ============================================================

const EXTERNAL_LOG_SHEET_ID = '12rlj7SR-Xofj8tdyu3atA9kUTLC44Y6Ja9cNfd2PUvw'; // 🔴 ต้องแก้ก่อนใช้งาน
const EXTERNAL_LOG_SHEET_NAME = 'Daily_Close_Log';
const EXTERNAL_LOG_COL = { DATE: 1, SYMBOL: 2, PRICE: 3, HIGH: 4, LOW: 5, CLOSE: 6 }; // 1-based ตามภาพ (A-F)

// จำนวนวันขั้นต่ำที่ต้องมีถึงจะคำนวณได้อย่างน่าเชื่อถือ (ยิ่งเยอะยิ่งนิ่ง)
const MIN_DAYS_EMA50 = 50;
const MIN_DAYS_MACD = 35; // ทางเทคนิคขั้นต่ำ 26+9 แต่แนะนำให้มีมากกว่านี้เพื่อความนิ่ง

// ══════════════════════════════════════════════════════════
// ดึงราคาปิดย้อนหลังของ ticker จาก Daily_Close_Log เรียงตามวันที่ (เก่า→ใหม่)
// คืน null ถ้าเปิดไฟล์ไม่ได้ หรือไม่พบ ticker เลย — ไม่ fabricate ข้อมูล
// ══════════════════════════════════════════════════════════
function _getClosesFromExternalLog(ticker) {
  try {
    if (!EXTERNAL_LOG_SHEET_ID || EXTERNAL_LOG_SHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
      throw new Error('ยังไม่ได้ตั้งค่า EXTERNAL_LOG_SHEET_ID — ใส่ Spreadsheet ID ของไฟล์ Daily_Close_Log ก่อน');
    }
    ticker = String(ticker || '').trim().toUpperCase();

    const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);
    const sheet = extSS.getSheetByName(EXTERNAL_LOG_SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีต "' + EXTERNAL_LOG_SHEET_NAME + '" ในไฟล์ภายนอก');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const matched = rows
      .filter(r => String(r[EXTERNAL_LOG_COL.SYMBOL - 1] || '').trim().toUpperCase() === ticker)
      .map(r => ({
        date: r[EXTERNAL_LOG_COL.DATE - 1],
        close: Number(r[EXTERNAL_LOG_COL.CLOSE - 1]),
        low: Number(r[EXTERNAL_LOG_COL.LOW - 1]),
        high: Number(r[EXTERNAL_LOG_COL.HIGH - 1])
      }))
      .filter(x => x.date && !isNaN(x.close))
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    return matched.length ? matched : null;
  } catch (e) {
    logError('_getClosesFromExternalLog', e);
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// คำนวณ EMA20/EMA50/MACD จากข้อมูลใน Daily_Close_Log
// คืน { available: true/false, ... } — available:false ถ้าข้อมูลไม่พอ
// (ไม่เดาค่า ไม่เติมข้อมูลลอยๆ — บอกตรงๆ ว่าขาดกี่วัน)
// ══════════════════════════════════════════════════════════
function getExtendedTechnicals(ticker) {
  try {
    const history = _getClosesFromExternalLog(ticker);
    if (!history) {
      return { available: false, reason: 'ไม่พบข้อมูลราคาย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log (อาจเป็นหุ้นที่ยังไม่ได้ถือ หรือเพิ่งเริ่มบันทึก)' };
    }

    const closes = history.map(h => h.close);
    const n = closes.length;

    const hasEMA50Data = n >= MIN_DAYS_EMA50;
    const hasMACDData = n >= MIN_DAYS_MACD;

    if (!hasEMA50Data && !hasMACDData) {
      return { available: false, reason: `มีข้อมูลราคาย้อนหลังแค่ ${n} วัน (ต้องการอย่างน้อย ${MIN_DAYS_MACD} วันสำหรับ MACD และ ${MIN_DAYS_EMA50} วันสำหรับ EMA50) — รอสะสมข้อมูลเพิ่ม` };
    }

    const result = { available: true, daysAvailable: n, currentPrice: closes[n - 1] };

    const ema20Arr = _extCalcEMA(closes, 20);
    result.ema20 = ema20Arr[n - 1];

    if (hasEMA50Data) {
      const ema50Arr = _extCalcEMA(closes, 50);
      result.ema50 = ema50Arr[n - 1];
      result.priceAboveEMA50 = closes[n - 1] > result.ema50;
      result.ema20AboveEMA50 = result.ema20 > result.ema50;
    } else {
      result.ema50Note = `ยังไม่พอ (มี ${n}/${MIN_DAYS_EMA50} วัน)`;
    }

    if (hasMACDData) {
      const macd = _extCalcMACD(closes);
      result.macdLine = macd.macdLine;
      result.signalLine = macd.signalLine;
      result.macdHistogram = macd.histogram;
      result.macdBullish = macd.macdLine > macd.signalLine;
    } else {
      result.macdNote = `ยังไม่พอ (มี ${n}/${MIN_DAYS_MACD} วัน)`;
    }

    return result;
  } catch (e) {
    logError('getExtendedTechnicals', e);
    return { available: false, reason: e.message };
  }
}

// ── EMA มาตรฐาน — เขียนแยกจากของเดิมในไฟล์อื่น (ตั้งชื่อ prefix _ext กันชนกัน
//    ตามบทเรียนจากปัญหา fetchYahooHistory ชื่อซ้ำก่อนหน้านี้) ──
function _extCalcEMA(closes, period) {
  const k = 2 / (period + 1);
  const ema = new Array(closes.length).fill(null);
  // เริ่มด้วย SMA ของ period แรกเป็นฐาน
  let sum = 0;
  for (let i = 0; i < period && i < closes.length; i++) sum += closes[i];
  if (closes.length < period) return ema;
  ema[period - 1] = sum / period;
  for (let i = period; i < closes.length; i++) {
    ema[i] = closes[i] * k + ema[i - 1] * (1 - k);
  }
  return ema;
}

// ── MACD = EMA12 - EMA26, Signal = EMA9 ของ MACD Line ──
function _extCalcMACD(closes) {
  const ema12 = _extCalcEMA(closes, 12);
  const ema26 = _extCalcEMA(closes, 26);
  const n = closes.length;

  const macdSeries = [];
  for (let i = 0; i < n; i++) {
    macdSeries.push((ema12[i] !== null && ema26[i] !== null) ? ema12[i] - ema26[i] : null);
  }
  const macdValid = macdSeries.filter(v => v !== null);
  const signalSeries = _extCalcEMA(macdValid, 9);

  const macdLine = macdValid[macdValid.length - 1];
  const signalLine = signalSeries[signalSeries.length - 1];

  return {
    macdLine, signalLine,
    histogram: (macdLine !== null && signalLine !== null) ? macdLine - signalLine : null
  };
}

// ══════════════════════════════════════════════════════════
// หาแนวรับ (Support) จาก Swing Low — จุดที่ราคา Low ต่ำสุดในช่วง ±WINDOW วัน
// ใช้ข้อมูลจาก Daily_Close_Log เดียวกับ getExtendedTechnicals (หุ้นที่ถืออยู่เท่านั้น)
// ══════════════════════════════════════════════════════════
const SWING_LOW_WINDOW = 3;     // เทียบกับ 3 วันก่อน-หลัง
const SWING_LOW_MIN_DAYS = 15;  // ต้องมีข้อมูลอย่างน้อยเท่านี้ถึงจะหา swing low ได้อย่างมีความหมาย
const SWING_LOW_MAX_RESULTS = 5; // เก็บแนวรับล่าสุดไว้กี่จุด

function findSupportLevels(ticker) {
  try {
    const history = _getClosesFromExternalLog(ticker);
    if (!history) {
      return { available: false, reason: 'ไม่พบข้อมูลราคาย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    if (history.length < SWING_LOW_MIN_DAYS) {
      return { available: false, reason: `มีข้อมูลแค่ ${history.length} วัน (ต้องการอย่างน้อย ${SWING_LOW_MIN_DAYS} วัน) — รอสะสมข้อมูลเพิ่ม` };
    }

    const lows = history.map(h => h.low);
    const currentPrice = history[history.length - 1].close;
    const n = lows.length;

    // ── หาจุด Swing Low: Low[i] ต้องต่ำสุดในช่วง [i-W, i+W] ──
    const swingLows = [];
    for (let i = SWING_LOW_WINDOW; i < n - SWING_LOW_WINDOW; i++) {
      const windowSlice = lows.slice(i - SWING_LOW_WINDOW, i + SWING_LOW_WINDOW + 1);
      const isLowestInWindow = lows[i] === Math.min(...windowSlice);
      if (isLowestInWindow && !isNaN(lows[i]) && lows[i] > 0) {
        swingLows.push({ date: history[i].date, price: lows[i] });
      }
    }

    if (!swingLows.length) {
      return { available: false, reason: 'ไม่พบจุด Swing Low ที่ชัดเจนในช่วงข้อมูลที่มี (ราคาอาจแกว่งตัวราบเรียบเกินไป)' };
    }

    // เรียงจากล่าสุด → เก่า แล้วตัดเก็บแค่ N จุดล่าสุด
    swingLows.sort((a, b) => new Date(b.date) - new Date(a.date));
    const recentSwingLows = swingLows.slice(0, SWING_LOW_MAX_RESULTS);

    // ── แนวรับหลัก = swing low ที่ราคาสูงสุดในบรรดาที่ยัง "ต่ำกว่าราคาปัจจุบัน" (ใกล้สุด/เกี่ยวข้องสุด) ──
    const belowCurrent = recentSwingLows.filter(s => s.price < currentPrice);
    let nearestSupport = null;
    if (belowCurrent.length) {
      nearestSupport = belowCurrent.reduce((max, s) => s.price > max.price ? s : max, belowCurrent[0]);
    }

    return {
      available: true,
      currentPrice,
      nearestSupport, // { date, price } หรือ null ถ้าราคาปัจจุบันต่ำกว่า swing low ทุกจุดที่เจอ (หลุดแนวรับหมดแล้ว)
      allSwingLows: recentSwingLows,
      daysAnalyzed: n
    };
  } catch (e) {
    logError('findSupportLevels', e);
    return { available: false, reason: e.message };
  }
}
// ── ทดสอบการเชื่อมต่อไฟล์ภายนอก — รันเองใน Apps Script Editor เพื่อเช็คว่าตั้งค่าถูกไหม ──
function testExternalLogConnection() {
  const test = _getClosesFromExternalLog('NVDA'); // เปลี่ยนเป็น ticker ที่มีจริงในไฟล์คุณ
  if (test) {
    Logger.log('✅ เชื่อมต่อสำเร็จ พบข้อมูล ' + test.length + ' วัน');
    Logger.log('วันแรก: ' + test[0].date + ' close=' + test[0].close);
    Logger.log('วันล่าสุด: ' + test[test.length - 1].date + ' close=' + test[test.length - 1].close);
  } else {
    Logger.log('❌ เชื่อมต่อไม่สำเร็จ หรือไม่พบ ticker นี้ — เช็ค EXTERNAL_LOG_SHEET_ID');
  }
}

// ══════════════════════════════════════════════════════════
// แนวรับ (Support Levels) — จาก High/Low รายวันใน Daily_Close_Log
// ใช้ 2 วิธีร่วมกัน: Rolling Low (ค่าต่ำสุดใน N วัน) + Swing Low (จุดต่ำสุดเฉพาะที่จริง)
// ตรงกับแนวคิด "Swing Low" ที่ระบุไว้เป็นข้อมูลกลางของทั้ง 2 สาย (Fast/Portfolio)
// ══════════════════════════════════════════════════════════

const MIN_DAYS_SUPPORT = 20; // ต้องมีอย่างน้อย 20 วันถึงจะเริ่มหา rolling low ได้อย่างมีความหมาย
//const SWING_LOW_WINDOW = 3;  // จุดจะถือว่าเป็น swing low ถ้าต่ำสุดในช่วง ±3 วัน

// ── ดึง Date/High/Low/Close ของ ticker จาก Daily_Close_Log เรียงตามวันที่ (เก่า→ใหม่) ──
function _getOHLCFromExternalLog(ticker) {
  try {
    if (!EXTERNAL_LOG_SHEET_ID || EXTERNAL_LOG_SHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
      throw new Error('ยังไม่ได้ตั้งค่า EXTERNAL_LOG_SHEET_ID');
    }
    ticker = String(ticker || '').trim().toUpperCase();

    const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);
    const sheet = extSS.getSheetByName(EXTERNAL_LOG_SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีต "' + EXTERNAL_LOG_SHEET_NAME + '"');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const matched = rows
      .filter(r => String(r[EXTERNAL_LOG_COL.SYMBOL - 1] || '').trim().toUpperCase() === ticker)
      .map(r => ({
        date: r[EXTERNAL_LOG_COL.DATE - 1],
        high: Number(r[EXTERNAL_LOG_COL.HIGH - 1]),
        low: Number(r[EXTERNAL_LOG_COL.LOW - 1]),
        close: Number(r[EXTERNAL_LOG_COL.CLOSE - 1])
      }))
      .filter(x => x.date && !isNaN(x.low) && !isNaN(x.close))
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    return matched.length ? matched : null;
  } catch (e) {
    logError('_getOHLCFromExternalLog', e);
    return null;
  }
}

// ── หาจุด Swing Low: วันที่ Low ต่ำกว่า Low ของทุกวันในช่วง ±window รอบข้าง ──
function _findSwingLows(ohlc, window) {
  const swings = [];
  for (let i = window; i < ohlc.length - window; i++) {
    const current = ohlc[i].low;
    let isSwingLow = true;
    for (let j = i - window; j <= i + window; j++) {
      if (j === i) continue;
      if (ohlc[j].low < current) { isSwingLow = false; break; }
    }
    if (isSwingLow) swings.push({ date: ohlc[i].date, low: current });
  }
  return swings;
}

// ══════════════════════════════════════════════════════════
// MAIN: หาแนวรับของหุ้นตัวนี้ — คืน available:false พร้อมเหตุผล ถ้าข้อมูลไม่พอ (ไม่เดา)
// ══════════════════════════════════════════════════════════
function getSupportLevels(ticker) {
  try {
    const ohlc = _getOHLCFromExternalLog(ticker);
    if (!ohlc) {
      return { available: false, reason: 'ไม่พบข้อมูล High/Low ย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    const n = ohlc.length;
    if (n < MIN_DAYS_SUPPORT) {
      return { available: false, reason: `มีข้อมูลแค่ ${n} วัน (ต้องการอย่างน้อย ${MIN_DAYS_SUPPORT} วัน) — รอสะสมข้อมูลเพิ่ม` };
    }

    const currentPrice = ohlc[n - 1].close;

    // ── Rolling Low: ค่าต่ำสุดใน 20/50 วันล่าสุด ──
    const last20 = ohlc.slice(Math.max(0, n - 20));
    const last50 = ohlc.slice(Math.max(0, n - 50));
    const rolling20Low = Math.min(...last20.map(x => x.low));
    const rolling50Low = n >= 50 ? Math.min(...last50.map(x => x.low)) : null;

    // ── Swing Low: จุดต่ำสุดเฉพาะที่จริง เอา 3 จุดล่าสุด ──
    const swings = _findSwingLows(ohlc, SWING_LOW_WINDOW)
      .sort((a, b) => new Date(b.date) - new Date(a.date))
      .slice(0, 3)
      .map(s => ({ date: Utilities.formatDate(new Date(s.date), 'Asia/Bangkok', 'dd/MM/yyyy'), low: s.low }));

    // ── แนวรับที่ "ใช้งานได้จริงที่สุด" — เอาค่าที่ใกล้ราคาปัจจุบันที่สุด (แต่ต้องต่ำกว่าราคาปัจจุบัน) ──
    const candidates = [rolling20Low, rolling50Low, ...swings.map(s => s.low)]
      .filter(v => v !== null && v < currentPrice);
    const nearestSupport = candidates.length ? Math.max(...candidates) : null;
    const distancePct = nearestSupport ? ((currentPrice - nearestSupport) / currentPrice) * 100 : null;

    return {
      available: true, daysAvailable: n, currentPrice,
      rolling20Low, rolling50Low, swings,
      nearestSupport, distancePct
    };
  } catch (e) {
    logError('getSupportLevels', e);
    return { available: false, reason: e.message };
  }
}



// ══════════════════════════════════════════════════════════
// WATCHLIST RISK METRICS — ATR (14 วัน) + MOS ประมาณการจาก P/E เทียบค่าเฉลี่ย
// (แก้ใหม่: ใช้ _getOHLCFromExternalLog() เดิม — เปิดไฟล์ภายนอกถูกต้อง +
//  ใช้ตำแหน่งคอลัมน์ EXTERNAL_LOG_COL แทนชื่อหัวตาราง)
// ══════════════════════════════════════════════════════════
function getWatchlistRiskMetrics(ticker, market) {
  try {
    const ohlc = _getOHLCFromExternalLog(ticker);
    if (!ohlc) {
      return { success: false, error: 'ไม่พบข้อมูล High/Low ย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    if (ohlc.length < 15) {
      return { success: false, error: `มีข้อมูลราคาแค่ ${ohlc.length} วัน (ต้องการอย่างน้อย 15 วันสำหรับ ATR)` };
    }

    // ── ATR (14 วัน, ค่าเฉลี่ยของ True Range) ──
    const trList = [];
    for (let i = 1; i < ohlc.length; i++) {
      const high = ohlc[i].high;
      const low = ohlc[i].low;
      const prevClose = ohlc[i - 1].close;
      const tr = Math.max(high - low, Math.abs(high - prevClose), Math.abs(low - prevClose));
      trList.push(tr);
    }
    const last14TR = trList.slice(-14);
    const atr = last14TR.reduce((s, v) => s + v, 0) / last14TR.length;
    const currentPrice = ohlc[ohlc.length - 1].close;
    const atrPct = currentPrice > 0 ? (atr / currentPrice) * 100 : null;

    // ── MOS ประมาณการจาก P/E ──
    const pe = _getPEComparisonForWatchlist(ticker, market, ohlc);

      return {
      success: true,
      atr, atrPct,
      currentPE: pe.currentPE,
      avgPE: pe.avgPE,
      mosPct: pe.mosPct,
      mosPeOnlyPct: pe.mosPeOnlyPct,
      pegRatio: pe.pegRatio,
      pegSource: pe.pegSource,
      growthPct: pe.growthPct,
      mosPegPct: pe.mosPegPct,
      peNote: pe.note
    };

  } catch (e) {
    logError('getWatchlistRiskMetrics', e);
    return { success: false, error: e.message };
  }
}



















// ══════════════════════════════════════════════════════════
// Yahoo crumb + cookie — จำเป็นสำหรับ endpoint quoteSummary (fundamentals)
// ตั้งแต่ Yahoo เริ่มบังคับ auth แบบนี้ (v8/finance/chart ยังไม่ต้องใช้)
// แคชไว้ 50 นาที กันขอใหม่ทุกครั้ง (ช้า + เสี่ยงโดน rate limit)
// ══════════════════════════════════════════════════════════
function _getYahooCrumbAndCookie() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('yahoo_crumb_cookie_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* แคชเสีย ขอใหม่ */ }
  }

  // ── Step 1: ขอ cookie จาก Yahoo ──
  const cookieResp = UrlFetchApp.fetch('https://fc.yahoo.com', {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });
  const headers = cookieResp.getAllHeaders();
  let setCookie = headers['Set-Cookie'] || headers['set-cookie'];
  if (!setCookie) throw new Error('ไม่ได้รับ Set-Cookie จาก Yahoo (fc.yahoo.com)');
  if (!Array.isArray(setCookie)) setCookie = [setCookie];

  const cookieStr = setCookie
    .map(c => String(c).split(';')[0]) // เอาแค่ name=value ตัดพวก Domain/Path/Expires ทิ้ง
    .join('; ');

  if (!cookieStr) throw new Error('แปลง cookie จาก Yahoo ไม่สำเร็จ');

  // ── Step 2: ใช้ cookie ขอ crumb ──
  const crumbResp = UrlFetchApp.fetch('https://query1.finance.yahoo.com/v1/test/getcrumb', {
    muteHttpExceptions: true,
    headers: { 'Cookie': cookieStr, 'User-Agent': 'Mozilla/5.0' }
  });
  const crumb = crumbResp.getContentText().trim();
  if (!crumb || crumb.includes('<html') || crumb.length > 100) {
    throw new Error('ขอ crumb ไม่สำเร็จ (HTTP ' + crumbResp.getResponseCode() + '): ' + crumb.slice(0, 100));
  }

  const result = { cookie: cookieStr, crumb };
  cache.put('yahoo_crumb_cookie_v1', JSON.stringify(result), 50 * 60); // แคช 50 นาที
  return result;
}


// ══════════════════════════════════════════════════════════
// เฟส 3 — Quote Extras: H/L วันนี้, ปริมาณ, มูลค่าตลาด, After-Hours, 52wk
// ⚠️ ส่วน H/L/Volume/After-hours ใช้ chart endpoint (ไม่ต้อง auth เร็ว)
//    ส่วน marketCap ต้องใช้ quoteSummary + crumb — cache ไว้ 15 นาทีต่อ symbol
//    กันยิง crumb ซ้ำถี่ๆ (ของหนักสุด/เสี่ยง rate limit สุด)
// ══════════════════════════════════════════════════════════
function getStockQuoteExtras(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const symbol = (market === 'TH') ? (ticker + '.BK') : ticker;
    const cache = CacheService.getScriptCache();

    // ── ส่วนที่ 1: chart endpoint — ไม่ต้อง auth ──
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol)
      + '?includePrePost=true&interval=1d&range=1d';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) return { success: false, error: 'ดึงราคาไม่สำเร็จ (HTTP ' + res.getResponseCode() + ')' };

    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result || !result.meta) return { success: false, error: 'ไม่พบข้อมูลของ ' + symbol };

    const m = result.meta;
    const price = m.regularMarketPrice || 0;
    const prevClose = m.chartPreviousClose || m.previousClose || null;
    const changePct = (prevClose && price) ? ((price - prevClose) / prevClose) * 100 : null;
    const hasPostMarket = m.postMarketPrice !== undefined && m.postMarketPrice !== null
      && m.marketState && m.marketState !== 'REGULAR';

    const out = {
      success: true, ticker, symbol, price, changePct,
      dayHigh: m.regularMarketDayHigh || null,
      dayLow:  m.regularMarketDayLow  || null,
      volume:  m.regularMarketVolume  || null,
      fiftyTwoWeekHigh: m.fiftyTwoWeekHigh || null,
      fiftyTwoWeekLow:  m.fiftyTwoWeekLow  || null,
      exchangeName: m.fullExchangeName || m.exchangeName || null,
      marketState: m.marketState || null,
      regularMarketTime: m.regularMarketTime
        ? Utilities.formatDate(new Date(m.regularMarketTime * 1000), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
        : null,
      afterHours: hasPostMarket ? {
        price: m.postMarketPrice,
        changePct: price ? ((m.postMarketPrice - price) / price) * 100 : null,
        time: m.postMarketTime
          ? Utilities.formatDate(new Date(m.postMarketTime * 1000), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
          : null
      } : null,
      marketCap: null
    };

    // ── ส่วนที่ 2: marketCap ผ่าน quoteSummary + crumb — เช็ค cache ก่อนเสมอ ──
    const mcCacheKey = 'marketcap_v1_' + symbol;
    const mcCached = cache.get(mcCacheKey);
    if (mcCached) {
      try {
        const mc = JSON.parse(mcCached);
        out.marketCap = mc.marketCap;
        if (!out.exchangeName) out.exchangeName = mc.exchangeName;
      } catch (e) { /* cache เสีย ข้ามไปขอใหม่ */ }
    } else {
      try {
        const auth = _getYahooCrumbAndCookie(); // reuse ของเดิมในไฟล์นี้
        const qsUrl = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/' + encodeURIComponent(symbol)
          + '?modules=price&crumb=' + encodeURIComponent(auth.crumb);
        let qsResp = UrlFetchApp.fetch(qsUrl, { muteHttpExceptions: true, headers: { 'Cookie': auth.cookie, 'User-Agent': 'Mozilla/5.0' } });

        if (qsResp.getResponseCode() === 401) {
          cache.remove('yahoo_crumb_cookie_v1');
          const auth2 = _getYahooCrumbAndCookie();
          const qsUrl2 = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/' + encodeURIComponent(symbol)
            + '?modules=price&crumb=' + encodeURIComponent(auth2.crumb);
          qsResp = UrlFetchApp.fetch(qsUrl2, { muteHttpExceptions: true, headers: { 'Cookie': auth2.cookie, 'User-Agent': 'Mozilla/5.0' } });
        }

        if (qsResp.getResponseCode() === 200) {
          const qsJson = JSON.parse(qsResp.getContentText());
          const qsResult = qsJson.quoteSummary && qsJson.quoteSummary.result && qsJson.quoteSummary.result[0];
          const priceModule = qsResult && qsResult.price;
          if (priceModule) {
            const mc = priceModule.marketCap ? priceModule.marketCap.raw : null;
            const ex = priceModule.exchangeName || null;
            out.marketCap = mc;
            if (!out.exchangeName) out.exchangeName = ex;
            cache.put(mcCacheKey, JSON.stringify({ marketCap: mc, exchangeName: ex }), 15 * 60); // 15 นาที
          }
        }
      } catch (authErr) {
        // ขอ marketCap ไม่สำเร็จ — ยังคืนข้อมูลส่วนอื่นได้ตามปกติ ไม่ให้ทั้งฟังก์ชันพัง
        logError('getStockQuoteExtras:marketCap', authErr);
      }
    }

    return out;
  } catch (e) {
    logError('getStockQuoteExtras', e);
    return { success: false, error: e.message };
  }
}








function _getPEComparisonForWatchlist(ticker, market, ohlc) {
  try {
    const symbol = market === 'TH' ? (ticker + '.BK') : ticker;

    let auth;
    try {
      auth = _getYahooCrumbAndCookie();
    } catch (authErr) {
      logError('_getPEComparisonForWatchlist:auth', authErr);
      return { currentPE: null, avgPE: null, mosPct: null, note: 'ขอสิทธิ์เข้าถึง Yahoo Finance ไม่สำเร็จ: ' + authErr.message };
    }

    const url = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/' + encodeURIComponent(symbol)
      + '?modules=summaryDetail,defaultKeyStatistics&crumb=' + encodeURIComponent(auth.crumb);

    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Cookie': auth.cookie, 'User-Agent': 'Mozilla/5.0' }
    });
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code === 401) {
      // crumb/cookie หมดอายุ — ล้างแคชแล้วลองใหม่อีก 1 ครั้ง
      CacheService.getScriptCache().remove('yahoo_crumb_cookie_v1');
      logError('_getPEComparisonForWatchlist:401_retry', new Error('crumb หมดอายุ กำลังขอใหม่'));
      const auth2 = _getYahooCrumbAndCookie();
      const url2 = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/' + encodeURIComponent(symbol)
        + '?modules=summaryDetail,defaultKeyStatistics&crumb=' + encodeURIComponent(auth2.crumb);
      const resp2 = UrlFetchApp.fetch(url2, {
        muteHttpExceptions: true,
        headers: { 'Cookie': auth2.cookie, 'User-Agent': 'Mozilla/5.0' }
      });
      return _parsePEResponse(resp2.getResponseCode(), resp2.getContentText(), symbol, ohlc);
    }

    return _parsePEResponse(code, text, symbol, ohlc);
  } catch (e) {
    logError('_getPEComparisonForWatchlist', e);
    return { currentPE: null, avgPE: null, mosPct: null, note: 'ดึงข้อมูล P/E ไม่สำเร็จ: ' + e.message };
  }
}

function _parsePEResponse(code, text, symbol, ohlc) {
  if (code !== 200) {
    return { currentPE: null, avgPE: null, mosPct: null, note: `ดึงข้อมูล P/E ไม่สำเร็จ (HTTP ${code}) — ${text.slice(0, 150)}` };
  }

  let json;
  try { json = JSON.parse(text); } catch (e) {
    return { currentPE: null, avgPE: null, mosPct: null, note: 'ตอบกลับจาก Yahoo ไม่ใช่ JSON ที่ถูกต้อง' };
  }

  const result = json.quoteSummary && json.quoteSummary.result && json.quoteSummary.result[0];
  if (!result) {
    const errMsg = json.quoteSummary && json.quoteSummary.error && json.quoteSummary.error.description;
    return { currentPE: null, avgPE: null, mosPct: null, note: 'ไม่พบข้อมูลของ ' + symbol + (errMsg ? ' (' + errMsg + ')' : '') };
  }

  const dks = result.defaultKeyStatistics || {};
  const summaryDetail = result.summaryDetail || {};

  const currentPE = summaryDetail.trailingPE && summaryDetail.trailingPE.raw;
  const eps = dks.trailingEps && dks.trailingEps.raw;
  const currentPrice = ohlc[ohlc.length - 1].close;

  if (!currentPE || !eps) {
    return { currentPE: currentPE || null, avgPE: null, mosPct: null, note: 'ไม่พบข้อมูล P/E หรือ EPS ของ ' + symbol + ' (อาจเป็น ETF, กองทุน, หรือบริษัทที่ยังไม่มีกำไร)' };
  }

  // ── วิธีที่ 1: Trailing P/E เทียบค่าเฉลี่ยย้อนหลัง (ของเดิม) ──
  const historicalPE = ohlc.map(x => x.close / eps).filter(v => isFinite(v) && v > 0);
  if (!historicalPE.length) {
    return { currentPE, avgPE: null, mosPct: null, note: 'คำนวณ P/E ย้อนหลังไม่ได้' };
  }
  const avgPE = historicalPE.reduce((s, v) => s + v, 0) / historicalPE.length;
  const mosPeOnlyPct = avgPE > 0 ? ((avgPE - currentPE) / avgPE) * 100 : null;

  // ── วิธีที่ 2: PEG Ratio (รวม Growth Rate) ──
  // Yahoo ให้ pegRatio มาตรงๆ (อิง 5-year expected earnings growth) — แม่นกว่าคำนวณเอง
  // ถ้าไม่มี (พบบ่อยในหุ้นเล็ก/หุ้นไทยบางตัว) fallback ไปคำนวณเองจาก earningsQuarterlyGrowth (QoQ)
  let pegRatio = dks.pegRatio && dks.pegRatio.raw;
  let pegSource = null;
  let growthPct = null;

  if (pegRatio && isFinite(pegRatio) && pegRatio > 0) {
    pegSource = 'yahoo'; // ใช้ 5-year expected growth ของ Yahoo
    growthPct = currentPE / pegRatio; // ย้อนกลับมาโชว์ % growth ที่ Yahoo ใช้คำนวณ ให้ผู้ใช้เห็นที่มา
  } else {
    const eqg = dks.earningsQuarterlyGrowth && dks.earningsQuarterlyGrowth.raw;
    if (eqg && isFinite(eqg) && eqg > 0) {
      growthPct = eqg * 100;
      pegRatio = currentPE / growthPct;
      pegSource = 'computed'; // ⚠️ QoQ growth ระยะสั้น ไม่แม่นเท่า 5-year expected ของ Yahoo
    }
  }

  let mosPegPct = null, fairValuePEG = null;
  if (pegRatio && growthPct && growthPct > 0) {
    const fairPE_peg = growthPct; // สมมติ PEG เป้าหมาย = 1.0 (Peter Lynch: ราคายุติธรรมเมื่อ P/E ≈ Growth%)
    fairValuePEG = eps * fairPE_peg;
    mosPegPct = fairValuePEG > 0 ? ((fairValuePEG - currentPrice) / fairValuePEG) * 100 : null;
  }

  // ── ผสม 2 วิธี 50/50 — ถ้าไม่มี PEG (หา growth ไม่ได้เลย) fallback ไปใช้ P/E อย่างเดียวเหมือนเดิม ──
  const mosCombinedPct = (mosPeOnlyPct !== null && mosPegPct !== null)
    ? (mosPeOnlyPct + mosPegPct) / 2
    : mosPeOnlyPct;

  return {
    currentPE, avgPE,
    mosPct: mosCombinedPct,   // ← ค่าหลักที่ใช้แสดงผล (ตอนนี้เป็นค่าผสมแล้ว ไม่ต้องแก้จุดเรียกใช้เดิม)
    mosPeOnlyPct,
    pegRatio, pegSource, growthPct,
    mosPegPct,
    note: null
  };
}


// ══════════════════════════════════════════════════════════
// แนวต้าน (Resistance Levels) — กลับด้านจากแนวรับ ใช้ High แทน Low
// เพิ่ม All-Time High จากชีต Highest_Close_Summary (อยู่ไฟล์เดียวกับ Daily_Close_Log)
// เป็นแนวต้านที่สำคัญกว่า rolling high ธรรมดา — จุดสูงสุดที่เคยบันทึกไว้ทั้งหมด
// ══════════════════════════════════════════════════════════

const HIGHEST_CLOSE_SHEET_NAME = 'Highest_Close_Summary';
const HIGHEST_CLOSE_COL = { SYMBOL: 1, HIGHEST_CLOSE: 2, DATE: 3 }; // A, B, C ตามภาพจริง

// ── หาจุด Swing High: วันที่ High สูงกว่า High ของทุกวันในช่วง ±window รอบข้าง ──
function _findSwingHighs(ohlc, window) {
  const swings = [];
  for (let i = window; i < ohlc.length - window; i++) {
    const current = ohlc[i].high;
    let isSwingHigh = true;
    for (let j = i - window; j <= i + window; j++) {
      if (j === i) continue;
      if (ohlc[j].high > current) { isSwingHigh = false; break; }
    }
    if (isSwingHigh) swings.push({ date: ohlc[i].date, high: current });
  }
  return swings;
}

// ── ดึงจุดสูงสุดตลอดกาลจากชีต Highest_Close_Summary — คืน null ถ้าไม่พบ ไม่เดา ──
function _getAllTimeHighFromSummary(ticker) {
  try {
    if (!EXTERNAL_LOG_SHEET_ID || EXTERNAL_LOG_SHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') return null;
    ticker = String(ticker || '').trim().toUpperCase();

    const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);
    const sheet = extSS.getSheetByName(HIGHEST_CLOSE_SHEET_NAME);
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const found = rows.find(r => String(r[HIGHEST_CLOSE_COL.SYMBOL - 1] || '').trim().toUpperCase() === ticker);
    if (!found) return null;

    const highestClose = Number(found[HIGHEST_CLOSE_COL.HIGHEST_CLOSE - 1]);
    const date = found[HIGHEST_CLOSE_COL.DATE - 1];
    if (isNaN(highestClose)) return null;

    return {
      price: highestClose,
      date: date instanceof Date ? Utilities.formatDate(date, 'Asia/Bangkok', 'dd/MM/yyyy') : String(date)
    };
  } catch (e) {
    logError('_getAllTimeHighFromSummary', e);
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// MAIN: หาแนวต้านของหุ้นตัวนี้ — คืน available:false พร้อมเหตุผล ถ้าข้อมูลไม่พอ (ไม่เดา)
// ══════════════════════════════════════════════════════════
function getResistanceLevels(ticker) {
  try {
    const ohlc = _getOHLCFromExternalLog(ticker);
    if (!ohlc) {
      return { available: false, reason: 'ไม่พบข้อมูล High/Low ย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    const n = ohlc.length;
    if (n < MIN_DAYS_SUPPORT) {
      return { available: false, reason: `มีข้อมูลแค่ ${n} วัน (ต้องการอย่างน้อย ${MIN_DAYS_SUPPORT} วัน) — รอสะสมข้อมูลเพิ่ม` };
    }

    const currentPrice = ohlc[n - 1].close;

    // ── Rolling High: ค่าสูงสุดใน 20/50 วันล่าสุด ──
    const last20 = ohlc.slice(Math.max(0, n - 20));
    const last50 = ohlc.slice(Math.max(0, n - 50));
    const rolling20High = Math.max(...last20.map(x => x.high));
    const rolling50High = n >= 50 ? Math.max(...last50.map(x => x.high)) : null;

    // ── Swing High: จุดสูงสุดเฉพาะที่จริง เอา 3 จุดล่าสุด ──
    const swings = _findSwingHighs(ohlc, SWING_LOW_WINDOW)
      .sort((a, b) => new Date(b.date) - new Date(a.date))
      .slice(0, 3)
      .map(s => ({ date: Utilities.formatDate(new Date(s.date), 'Asia/Bangkok', 'dd/MM/yyyy'), high: s.high }));

    // ── All-Time High จาก Highest_Close_Summary — แนวต้านสำคัญที่สุด ──
    const allTimeHigh = _getAllTimeHighFromSummary(ticker);

    // ── แนวต้านที่ "ใช้งานได้จริงที่สุด" — เอาค่าที่ใกล้ราคาปัจจุบันที่สุด (แต่ต้องสูงกว่าราคาปัจจุบัน) ──
    const candidates = [rolling20High, rolling50High, ...swings.map(s => s.high), allTimeHigh ? allTimeHigh.price : null]
      .filter(v => v !== null && v > currentPrice);
    const nearestResistance = candidates.length ? Math.min(...candidates) : null;
    const distancePct = nearestResistance ? ((nearestResistance - currentPrice) / currentPrice) * 100 : null;

    return {
      available: true, daysAvailable: n, currentPrice,
      rolling20High, rolling50High, swings, allTimeHigh,
      nearestResistance, distancePct
    };
  } catch (e) {
    logError('getResistanceLevels', e);
    return { available: false, reason: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// ดึงราคาปิดย้อนหลังสำหรับวาดกราฟเส้นในหน้าวิเคราะห์หุ้นรายตัว
// reuse _getClosesFromExternalLog() ตัวเดิม — เอาแค่ 90 วันล่าสุดพอ (กันกราฟแน่นเกินไป)
// ══════════════════════════════════════════════════════════
function getPriceChartData(ticker) {
  try {
    const history = _getClosesFromExternalLog(ticker);
    if (!history || history.length < 2) {
      return { success: false, error: 'ไม่มีข้อมูลราคาย้อนหลังเพียงพอสำหรับวาดกราฟ' };
    }
    const recent = history.slice(-90); // 90 วันล่าสุด
    return {
      success: true,
      closes: recent.map(h => h.close),
      startDate: Utilities.formatDate(new Date(recent[0].date), 'Asia/Bangkok', 'dd/MM/yy'),
      endDate: Utilities.formatDate(new Date(recent[recent.length - 1].date), 'Asia/Bangkok', 'dd/MM/yy'),
      daysAvailable: recent.length
    };
  } catch (e) {
    logError('getPriceChartData', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// เฟส 4 — กราฟแท่งเทียนรายชั่วโมง + MA5/10/20 (สำหรับหน้า Stock Card)
// ใช้ chart endpoint แบบ interval=60m — ไม่ต้อง auth (เหมือน _wlFetchYahooQuote)
// ══════════════════════════════════════════════════════════
function getHourlyCandleData(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const symbol = (market === 'TH') ? (ticker + '.BK') : ticker;
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol)
      + '?interval=60m&range=5d';

    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) return { success: false, error: 'ดึงกราฟรายชั่วโมงไม่สำเร็จ (HTTP ' + res.getResponseCode() + ')' };

    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result || !result.timestamp || !result.indicators || !result.indicators.quote || !result.indicators.quote[0]) {
      return { success: false, error: 'ไม่พบข้อมูลกราฟรายชั่วโมงของ ' + symbol };
    }

    const ts = result.timestamp;
    const q = result.indicators.quote[0];
    const bars = [];
    for (let i = 0; i < ts.length; i++) {
      const o = q.open[i], h = q.high[i], l = q.low[i], c = q.close[i];
      if (o == null || h == null || l == null || c == null) continue; // ข้ามแท่งที่ตลาดปิด
      bars.push({ t: ts[i], o, h, l, c });
    }
    if (bars.length < 5) return { success: false, error: 'ข้อมูลรายชั่วโมงมีไม่พอสำหรับวาดกราฟ' };

    const recent = bars.slice(-60); // เอา 60 แท่งล่าสุด กันกราฟแน่นเกินไปบนมือถือ
    const closes = recent.map(b => b.c);

    const prevClose = result.meta.chartPreviousClose || result.meta.previousClose || closes[0];
    const lastClose = closes[closes.length - 1];
    const changePct = prevClose ? ((lastClose - prevClose) / prevClose) * 100 : null;

    return {
      success: true, ticker, symbol,
      bars: recent.map(b => ({
        time: Utilities.formatDate(new Date(b.t * 1000), 'Asia/Bangkok', 'dd/MM HH:mm'),
        o: b.o, h: b.h, l: b.l, c: b.c
      })),
      ma5:  _candleRollingMA(closes, 5),
      ma10: _candleRollingMA(closes, 10),
      ma20: _candleRollingMA(closes, 20),
      changePct, lastClose,
      timeframe: '1 ชม.'
    };
  } catch (e) {
    logError('getHourlyCandleData', e);
    return { success: false, error: e.message };
  }
}

// ── SMA รายจุด (rolling) — คืน null ในช่วงต้นที่ข้อมูลยังไม่ครบ period ──
function _candleRollingMA(closes, period) {
  return closes.map((_, i) => {
    if (i < period - 1) return null;
    let sum = 0;
    for (let j = i - period + 1; j <= i; j++) sum += closes[j];
    return sum / period;
  });
}



// ══════════════════════════════════════════════════════════
// ADX (Average Directional Index) — วัดความแข็งแรงของเทรนด์ (ไม่บอกทิศทาง)
// ใช้ +DI/-DI ร่วมด้วยเพื่อบอกทิศทางที่เทรนด์แข็งแรงอยู่ (bullish/bearish)
// สูตร Wilder's Smoothing มาตรฐาน — ต้องการอย่างน้อย ~35 วัน (14 DM/TR + 14 DX smoothing + buffer)
// ══════════════════════════════════════════════════════════
const ADX_PERIOD = 14;
const ADX_MIN_DAYS = 35;

function getADXData(ticker) {
  try {
    const ohlc = _getOHLCFromExternalLog(ticker); // ใช้ตัวเดียวกับ Support/Resistance
    if (!ohlc) {
      return { available: false, reason: 'ไม่พบข้อมูล High/Low ย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    const n = ohlc.length;
    if (n < ADX_MIN_DAYS) {
      return { available: false, reason: `มีข้อมูลแค่ ${n} วัน (ต้องการอย่างน้อย ${ADX_MIN_DAYS} วันสำหรับ ADX) — รอสะสมข้อมูลเพิ่ม` };
    }

    const result = _extCalcADX(ohlc, ADX_PERIOD);
    if (!result) {
      return { available: false, reason: 'คำนวณ ADX ไม่สำเร็จ (ข้อมูลไม่ต่อเนื่องพอ)' };
    }

    let strength, strengthClass;
    if (result.adx < 20) { strength = 'ไม่มีเทรนด์ชัดเจน / Sideway'; strengthClass = 'warn'; }
    else if (result.adx < 25) { strength = 'เทรนด์เริ่มก่อตัว'; strengthClass = 'warn'; }
    else if (result.adx < 40) { strength = 'เทรนด์แข็งแรง'; strengthClass = 'safe'; }
    else { strength = 'เทรนด์แข็งแรงมาก'; strengthClass = 'safe'; }

    const direction = result.plusDI > result.minusDI ? 'bullish' : 'bearish';
    const directionLabel = direction === 'bullish' ? '📈 ฝั่งขาขึ้นแข็งแรงกว่า (+DI > -DI)' : '📉 ฝั่งขาลงแข็งแรงกว่า (-DI > +DI)';

    return {
      available: true,
      adx: _taxRound(result.adx, 1),
      plusDI: _taxRound(result.plusDI, 1),
      minusDI: _taxRound(result.minusDI, 1),
      strength, strengthClass,
      direction, directionLabel,
      daysAvailable: n
    };
  } catch (e) {
    logError('getADXData', e);
    return { available: false, reason: e.message };
  }
}

// ── Wilder's Smoothing มาตรฐาน — คืน { adx, plusDI, minusDI } ของค่าล่าสุด หรือ null ถ้าคำนวณไม่ได้ ──
function _extCalcADX(ohlc, period) {
  const n = ohlc.length;
  const tr = [], plusDM = [], minusDM = [];

  for (let i = 1; i < n; i++) {
    const upMove = ohlc[i].high - ohlc[i - 1].high;
    const downMove = ohlc[i - 1].low - ohlc[i].low;
    plusDM.push((upMove > downMove && upMove > 0) ? upMove : 0);
    minusDM.push((downMove > upMove && downMove > 0) ? downMove : 0);
    tr.push(Math.max(
      ohlc[i].high - ohlc[i].low,
      Math.abs(ohlc[i].high - ohlc[i - 1].close),
      Math.abs(ohlc[i].low - ohlc[i - 1].close)
    ));
  }

  if (tr.length < period * 2) return null;

  // ── ค่าเริ่มต้น: ผลรวม period แรก ──
  let trSum = tr.slice(0, period).reduce((s, v) => s + v, 0);
  let plusDMSum = plusDM.slice(0, period).reduce((s, v) => s + v, 0);
  let minusDMSum = minusDM.slice(0, period).reduce((s, v) => s + v, 0);

  const dxList = [];
  for (let i = period; i < tr.length; i++) {
    // Wilder's smoothing: ค่าใหม่ = ค่าเดิม - (ค่าเดิม/period) + ค่าล่าสุด
    trSum = trSum - (trSum / period) + tr[i];
    plusDMSum = plusDMSum - (plusDMSum / period) + plusDM[i];
    minusDMSum = minusDMSum - (minusDMSum / period) + minusDM[i];

    const plusDI = trSum > 0 ? (plusDMSum / trSum) * 100 : 0;
    const minusDI = trSum > 0 ? (minusDMSum / trSum) * 100 : 0;
    const diSum = plusDI + minusDI;
    const dx = diSum > 0 ? (Math.abs(plusDI - minusDI) / diSum) * 100 : 0;
    dxList.push({ dx, plusDI, minusDI });
  }

  if (dxList.length < period) return null;

  // ── ADX แรก = ค่าเฉลี่ยของ DX ชุดแรก (period ตัว) แล้ว smooth ต่อแบบ Wilder's ──
  let adx = dxList.slice(0, period).reduce((s, v) => s + v.dx, 0) / period;
  for (let i = period; i < dxList.length; i++) {
    adx = ((adx * (period - 1)) + dxList[i].dx) / period;
  }

  const last = dxList[dxList.length - 1];
  return { adx, plusDI: last.plusDI, minusDI: last.minusDI };
}

// ── กันเผื่อ _taxRound ยังไม่ถูกโหลดในไฟล์นี้ (คนละไฟล์กับ webapp_16) ──
//if (typeof _taxRound === 'undefined') {
//  function _taxRound(v, d) {
//    const f = Math.pow(10, d || 0);
//    return Math.round((v + Number.EPSILON) * f) / f;
//  }
//}



// ══════════════════════════════════════════════════════════
// MAIN: หา Pivot Point + R1-R3 / S1-S3 ของหุ้นตัวนี้
// สูตรมาตรฐาน (Classic Pivot Point) คำนวณจาก High/Low/Close ของ "วันล่าสุด" ใน Daily_Close_Log
// คืน available:false พร้อมเหตุผล ถ้าข้อมูลไม่พอ (ไม่เดา)
// ══════════════════════════════════════════════════════════
function getPivotLevels(ticker) {
  try {
    const ohlc = _getOHLCFromExternalLog(ticker);
    if (!ohlc) {
      return { available: false, reason: 'ไม่พบข้อมูล High/Low ย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    const n = ohlc.length;
    if (n < 1) {
      return { available: false, reason: 'ยังไม่มีข้อมูลราคาสำหรับคำนวณ Pivot Point' };
    }

    // ── ใช้ High/Low/Close ของวันล่าสุดที่มีข้อมูลครบ เป็นฐานคำนวณ ──
    const last = ohlc[n - 1];
    const high = last.high, low = last.low, close = last.close;
    if ([high, low, close].some(v => v === null || isNaN(v))) {
      return { available: false, reason: 'ข้อมูล High/Low/Close ของวันล่าสุดไม่ครบ — คำนวณ Pivot ไม่ได้' };
    }

    const pivot = (high + low + close) / 3;
    const r1 = 2 * pivot - low;
    const r2 = pivot + (high - low);
    const r3 = high + 2 * (pivot - low);
    const s1 = 2 * pivot - high;
    const s2 = pivot - (high - low);
    const s3 = low - 2 * (high - pivot);

    const currentPrice = close; // ราคาปิดล่าสุดที่ใช้เป็นฐาน (เทียบตำแหน่งราคาปัจจุบันกับระดับต่างๆ)

    // ── หาว่าราคาปัจจุบันอยู่โซนไหน (ระหว่าง S ไหนกับ R ไหน) — ใช้บอก UI ว่าจะไฮไลต์เส้นไหน ──
    const allLevels = [
      { key: 'r3', price: r3 }, { key: 'r2', price: r2 }, { key: 'r1', price: r1 },
      { key: 's1', price: s1 }, { key: 's2', price: s2 }, { key: 's3', price: s3 }
    ];
    const nearestResistanceLevel = allLevels.filter(l => l.price > currentPrice).sort((a, b) => a.price - b.price)[0] || null;
    const nearestSupportLevel = allLevels.filter(l => l.price < currentPrice).sort((a, b) => b.price - a.price)[0] || null;

    return {
      available: true,
      baseDate: last.date instanceof Date ? Utilities.formatDate(last.date, 'Asia/Bangkok', 'dd/MM/yyyy') : String(last.date),
      currentPrice,
      pivot,
      resistance: { r1, r2, r3 },
      support: { s1, s2, s3 },
      nearestResistanceLevel, // { key, price } หรือ null ถ้าราคาสูงกว่า R3 ไปแล้ว
      nearestSupportLevel     // { key, price } หรือ null ถ้าราคาต่ำกว่า S3 ไปแล้ว
    };
  } catch (e) {
    logError('getPivotLevels', e);
    return { available: false, reason: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// สร้างเหตุผล/คำเตือนจากตำแหน่งราคาเทียบกับ Pivot Point
// ใช้ร่วมกันทั้งหุ้นที่ถือ (webapp_03c) และ Watchlist (webapp_07)
// ══════════════════════════════════════════════════════════
function buildPivotKeyTakeaways(pivot, currentPrice) {
  const reasons = [];
  const warnings = [];
  if (!pivot || !pivot.available) return { reasons, warnings };

  const lv = v => (v === null || v === undefined || isNaN(v)) ? '-' : v.toFixed(2);

  if (!pivot.nearestResistanceLevel) {
    reasons.push('ราคาทะลุแนวต้าน R3 ขึ้นไปแล้ว (' + lv(pivot.resistance.r3) + ') — โมเมนตัมขาขึ้นแข็งแกร่ง');
  } else {
    const distR = ((pivot.nearestResistanceLevel.price - currentPrice) / currentPrice) * 100;
    if (distR <= 2) {
      warnings.push('ราคาใกล้แนวต้าน ' + pivot.nearestResistanceLevel.key.toUpperCase() + ' ที่ ' + lv(pivot.nearestResistanceLevel.price) + ' (+' + distR.toFixed(1) + '%) — ระวังแรงขายทำกำไร');
    }
  }

  if (!pivot.nearestSupportLevel) {
    warnings.push('ราคาหลุดต่ำกว่าแนวรับ S3 (' + lv(pivot.support.s3) + ') แล้ว — ระวังแนวโน้มขาลงต่อเนื่อง');
  } else {
    const distS = ((currentPrice - pivot.nearestSupportLevel.price) / currentPrice) * 100;
    if (distS <= 2) {
      reasons.push('ราคาใกล้แนวรับ ' + pivot.nearestSupportLevel.key.toUpperCase() + ' ที่ ' + lv(pivot.nearestSupportLevel.price) + ' (-' + distS.toFixed(1) + '%) — เป็นโซนที่น่าสนใจเข้าซื้อ/เฝ้าดู');
    }
  }

  if (currentPrice >= pivot.pivot) {
    reasons.push('ราคายืนเหนือ Pivot (' + lv(pivot.pivot) + ') — แนวโน้มระยะสั้นเอนไปทางขาขึ้น');
  } else {
    warnings.push('ราคาต่ำกว่า Pivot (' + lv(pivot.pivot) + ') — แนวโน้มระยะสั้นเอนไปทางขาลง');
  }

  return { reasons, warnings };
}





/** ATR(14) จาก OHLC array — แยกออกมาเป็น helper ใช้ร่วมกับ getWatchlistRiskMetrics ได้ */
function _calcATR14(ohlc) {
  const clean = ohlc.filter(x =>
    Number.isFinite(x.high) && Number.isFinite(x.low) && Number.isFinite(x.close)
  );
  if (clean.length < 2) return 0;

  const trList = [];
  for (let i = 1; i < clean.length; i++) {
    const tr = Math.max(
      clean[i].high - clean[i].low,
      Math.abs(clean[i].high - clean[i - 1].close),
      Math.abs(clean[i].low - clean[i - 1].close)
    );
    trList.push(tr);
  }
  const last14 = trList.slice(-14);
  return last14.length ? last14.reduce((s, v) => s + v, 0) / last14.length : 0;
}


// ══════════════════════════════════════════════════════════
// TRAILING STOP สายถือ (Portfolio) — ล็อกกำไรแบบหลวม เน้นถือยาว ไม่ใช่คำสั่งขายเด็ดขาด
// ══════════════════════════════════════════════════════════
const PORTFOLIO_TRAIL_START_PROFIT_PCT = 10;  // เริ่มทำงานเมื่อกำไร ≥ 10%
const PORTFOLIO_TRAIL_ATR_MULTIPLIER   = 3.5; // หลวมกว่า Fast (1-2x) เพราะถือยาว

//ใช้ **Highest High** (ไม่ใช่ Highest Close), **returnPct จาก currentPrice เทียบ avgCost ตรงๆ** (ไม่พึ่ง unrealizedPL), **tier เรียงจาก 25%→50%→100%** (ตึงสุด→หลวมสุด, ทยอยลดความเสี่ยง), มี floor กันหลุดต่ำกว่าทุน, คำนวณจำนวนหุ้น/กำไรต่อ tier, และผูกกับ Execution Log

function getPortfolioTrailingStop(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const eff = getEffectiveRiskParams(ticker);

    const h = getHoldingsData();
    const arr = (market === 'TH') ? h.th : h.us;
    const holding = (arr || []).find(x => x.ticker === ticker);
    if (!holding) return { success: false, error: 'ไม่พบ ' + ticker + ' ใน Holdings' };

    const ohlcRaw = _getOHLCFromExternalLog(ticker);
    const ohlc = (ohlcRaw || []).filter(x =>
      x && Number.isFinite(x.high) && Number.isFinite(x.low) && Number.isFinite(x.close)
    );
    if (ohlc.length < 15) {
      return { success: false, error: 'ข้อมูลราคาย้อนหลังไม่พอ หรือมีข้อมูลผิดรูปแบบ (ต้องการอย่างน้อย 15 วันที่ใช้ได้จริง)' };
    }

    const firstBuyDate = _getFirstBuyDateForTicker(ticker, market);
    const sinceEntry = firstBuyDate ? ohlc.filter(x => new Date(x.date) >= firstBuyDate) : ohlc;
    if (sinceEntry.length < 2) {
      return { success: false, error: 'ถือมาน้อยเกินไป ยังไม่มีข้อมูลราคาพอคำนวณ' };
    }

    const atr = _calcATR14(ohlc);

    // ── Highest HIGH (ไม่ใช่ Close) ──
    const highestHighSinceEntry = Math.max(...sinceEntry.map(x => x.high));

    const lastBar = ohlc[ohlc.length - 1];
    const currentPrice = Number.isFinite(holding.priceNow) ? holding.priceNow : lastBar.close;

    const triggerPriceMap = {
      currentPrice: currentPrice,
      lastClose: lastBar.close,
      todayLow: lastBar.low
    };
    const triggerPrice = triggerPriceMap[eff.trailTriggerMode] ?? lastBar.close;

    // ── returnPct จาก currentPrice ตรงๆ ──
    const returnPct = holding.avgCost > 0
      ? ((currentPrice - holding.avgCost) / holding.avgCost) * 100
      : 0;

    const active = returnPct >= eff.portfolioTrailStartProfitPct;
    const floorPrice = holding.avgCost * (1 + eff.minProfitProtectPct / 100);

    // ── Tier เรียง 25% (ตึงสุด) → 50% → 100% (หลวมสุด) : ทยอยลดความเสี่ยง ──
    const tierDefs = [
      { key: 'sell25',  label: 'ขาย 25% ล็อกกำไรบางส่วน', multiplier: eff.portfolioTrailAtrX * 0.5, fraction: 0.25 },
      { key: 'sell50',  label: 'ขาย 50% ปล่อยครึ่งวิ่งต่อ', multiplier: eff.portfolioTrailAtrX * 1.0, fraction: 0.50 },
      { key: 'sell100', label: 'ขาย 100% (ออกทั้งหมด)',   multiplier: eff.portfolioTrailAtrX * 1.5, fraction: 1.00 }
    ];

    const executedTiers = _getExecutedTiers(ticker, market, highestHighSinceEntry, eff.trailResetPct);

    const tiers = tierDefs.map(t => {
      const rawStop = highestHighSinceEntry - atr * t.multiplier;
      const stopPrice = Math.max(rawStop, floorPrice);
      const floored = rawStop < floorPrice;
      const execInfo = executedTiers.get(t.key) || null;
      const executed = !!execInfo;
      const triggered = active && !executed && triggerPrice <= stopPrice;
      const distancePct = active && currentPrice > 0 ? ((currentPrice - stopPrice) / currentPrice) * 100 : null;

      const sharesToSell = holding.sharesRemain * t.fraction;
      const sharesRemainingAfter = holding.sharesRemain - sharesToSell;
      const estimatedSellValue = sharesToSell * stopPrice;
      const estimatedProfit = sharesToSell * (stopPrice - holding.avgCost);

      return { ...t, stopPrice, floored, triggered, executed, execInfo, distancePct,
        sharesToSell, sharesRemainingAfter, estimatedSellValue, estimatedProfit };
    });

    return {
      success: true, available: true, active,
      highestHighSinceEntry, currentPrice, triggerPrice,
      trailTriggerMode: eff.trailTriggerMode,
      atr, baseMultiplier: eff.portfolioTrailAtrX,
      startProfitPct: eff.portfolioTrailStartProfitPct,
      minProfitProtectPct: eff.minProfitProtectPct,
      floorPrice, avgCost: holding.avgCost, sharesRemain: holding.sharesRemain,
      returnPct, daysAvailable: sinceEntry.length,
      tiers
    };
  } catch (e) {
    logError('getPortfolioTrailingStop', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// ราคาปิดรายวัน 30 วันล่าสุด + % เปลี่ยนแปลงต่อวัน + % เปลี่ยนแปลงรวม 30 วัน
// ใช้เปิดใน Modal จากหน้า Stock Detail (Fast/Portfolio) และ Watchlist Analyze
// ══════════════════════════════════════════════════════════
function getDailyCloseHistory(ticker) {
  try {
    const history = _getClosesFromExternalLog(ticker);
    if (!history) {
      return { available: false, reason: 'ไม่พบข้อมูลราคาย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    if (history.length < 2) {
      return { available: false, reason: 'มีข้อมูลไม่พอ (ต้องการอย่างน้อย 2 วัน)' };
    }

    const last31 = history.slice(-31);
    const rows = [];
    for (let i = 1; i < last31.length; i++) {
      const prevClose = last31[i - 1].close;
      const close = last31[i].close;
      const high = last31[i].high;
      const low = last31[i].low; // ← เพิ่ม
      const changePct = prevClose > 0 ? ((close - prevClose) / prevClose) * 100 : null;
      rows.push({
        date: Utilities.formatDate(new Date(last31[i].date), 'Asia/Bangkok', 'dd/MM/yyyy'),
        close, high, low, changePct
      });
    }
    rows.reverse();

    const first = last31[0].close;
    const last = last31[last31.length - 1].close;
    const overallChangePct = first > 0 ? ((last - first) / first) * 100 : null;

    const highsInRange = rows.map(r => r.high).filter(v => isFinite(v) && v > 0);
    const periodHigh = highsInRange.length ? Math.max(...highsInRange) : null;
    const periodHighRow = periodHigh !== null ? rows.find(r => r.high === periodHigh) : null;

    // ── เพิ่ม: จุดต่ำสุดของช่วง ──
    const lowsInRange = rows.map(r => r.low).filter(v => isFinite(v) && v > 0);
    const periodLow = lowsInRange.length ? Math.min(...lowsInRange) : null;
    const periodLowRow = periodLow !== null ? rows.find(r => r.low === periodLow) : null;

    // ── เพิ่ม: ค่าเบี่ยงเบนมาตรฐานของ %เปลี่ยนแปลงรายวัน (ใช้ประมาณระยะย่อ "ปกติ") ──
    const validChanges = rows.map(r => r.changePct).filter(v => v !== null);
    let dailyVolatilityPct = null;
    if (validChanges.length >= 5) {
      const mean = validChanges.reduce((s, v) => s + v, 0) / validChanges.length;
      const variance = validChanges.reduce((s, v) => s + Math.pow(v - mean, 2), 0) / validChanges.length;
      dailyVolatilityPct = Math.sqrt(variance);
    }

    return {
      available: true,
      ticker,
      rows,
      daysCount: rows.length,
      overallChangePct,
      periodHigh,
      periodHighDate: periodHighRow ? periodHighRow.date : null,
      periodLow,                                                  // ← เพิ่ม
      periodLowDate: periodLowRow ? periodLowRow.date : null,      // ← เพิ่ม
      dailyVolatilityPct,                                         // ← เพิ่ม
      startDate: rows[rows.length - 1].date,
      endDate: rows[0].date
    };
  } catch (e) {
    logError('getDailyCloseHistory', e);
    return { available: false, reason: e.message };
  }
}




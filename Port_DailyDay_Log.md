## รหัส.gs
// ============================================================
//  รหัส.gs
//  backfillHistoricalPrices()
//  ดึงราคาย้อนหลัง 6 เดือน → Daily_Close_Log
//  รันครั้งเดียว แล้วรัน updateHighestClose() ต่อ
// ============================================================

function backfillHistoricalPrices() {
  const db       = getDB();
  const logSheet = getOrCreateSheet(db, "Daily_Close_Log");

  // สร้าง Header ถ้ายังไม่มี
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
    logSheet.setFrozenRows(1);
  }

  // ── ดึง Ticker จาก Backend_ATR ──────────────────────────
  const { us, th } = getTickers();
  Logger.log(`Backfill — US: [${us}] | TH: [${th}]`);

  // ── ช่วงวันที่: ย้อนหลัง 6 เดือน ──────────────────────
  const today    = new Date();
  const sixAgo   = new Date(today);
  sixAgo.setMonth(sixAgo.getMonth() - 6);

  const period1  = Math.floor(sixAgo.getTime() / 1000);   // Unix timestamp เริ่ม
  const period2  = Math.floor(today.getTime()  / 1000);   // Unix timestamp สิ้นสุด

  // ── ดึงข้อมูลที่มีอยู่แล้ว (เพื่อไม่ให้ซ้ำ) ──────────
  const existing = new Set();
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues()
      .forEach(r => existing.add(`${r[0]}_${r[1]}`));
  }

  // ── Helper: ดึง daily OHLC จาก Yahoo Finance ───────────
  function fetchHistory(symbol) {
    try {
      const url  = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}`
                 + `?interval=1d&period1=${period1}&period2=${period2}`;
      const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(res.getContentText());
      const result = json.chart.result[0];

      const timestamps = result.timestamp;
      const quotes     = result.indicators.quote[0];

      const rows = [];
      timestamps.forEach((ts, i) => {
        const close = quotes.close[i];
        const high  = quotes.high[i];
        const low   = quotes.low[i];
        if (!close) return;   // วันหยุดตลาด

        const d   = new Date(ts * 1000);
        const key = `${Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")}_${symbol.replace(".BK","")}`;
        if (existing.has(key)) return;  // มีแล้ว ข้าม

        rows.push([
          Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          symbol.replace(".BK", ""),
          close,
          high,
          low,
          close,
        ]);
      });
      return rows;
    } catch (e) {
      Logger.log(`fetchHistory error [${symbol}]: ${e}`);
      return [];
    }
  }

  // ── ดึงและรวบรวมทุก Symbol ──────────────────────────────
  let allRows = [];

  us.forEach(sym => {
    const rows = fetchHistory(sym);
    Logger.log(`${sym}: ${rows.length} วัน`);
    allRows = allRows.concat(rows);
    Utilities.sleep(500);  // กัน rate limit
  });

  th.forEach(sym => {
    const rows = fetchHistory(sym + ".BK");
    Logger.log(`${sym}: ${rows.length} วัน`);
    allRows = allRows.concat(rows);
    Utilities.sleep(500);
  });

  // ── เรียงตามวันที่ก่อนเขียน ────────────────────────────
  allRows.sort((a, b) => a[0].localeCompare(b[0]));

  // ── เขียนลง Daily_Close_Log ────────────────────────────
  if (allRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, allRows.length, 6)
            .setValues(allRows);
    Logger.log(`Backfill เสร็จ: บันทึก ${allRows.length} แถว`);
  } else {
    Logger.log("Backfill: ไม่มีข้อมูลใหม่");
  }

  // ── คำนวณ Highest Close ทันที ──────────────────────────
  updateHighestClose();

  // ── แจ้ง Telegram ──────────────────────────────────────
  const symbols = [...us, ...th].length;
  sendTelegram(
    `📥 <b>Backfill เสร็จแล้ว</b>\n` +
    `📅 ย้อนหลัง 6 เดือน\n` +
    `📊 ${symbols} Symbols | ${allRows.length} แถว\n` +
    `✅ Highest Close อัปเดตแล้ว`
  );

  SpreadsheetApp.getUi().alert(
    `✅ Backfill เสร็จแล้ว!\n\n` +
    `📅 ช่วง: ${Utilities.formatDate(sixAgo, Session.getScriptTimeZone(), "yyyy-MM-dd")} → วันนี้\n` +
    `📊 Symbols: ${symbols} ตัว\n` +
    `📋 บันทึก: ${allRows.length} แถว\n` +
    `💡 Highest Close คำนวณใหม่แล้ว`
  );
}



// ============================================================
//  cleanOldData()
//  ลบข้อมูลใน Daily_Close_Log ที่เก่ากว่า 60 วัน
//  แต่เก็บแถวที่เป็นค่าสูงสุด (Highest Close) ของแต่ละ Symbol ไว้
//  รันได้ด้วยมือ หรือตั้ง Trigger รายเดือน
// ============================================================

function cleanOldData() {
  const db       = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  if (!logSheet || logSheet.getLastRow() < 2) {
    writeLog("ข้อมูล", "cleanOldData: ไม่มีข้อมูลใน Daily_Close_Log");
    return;
  }

  writeLog("ข้อมูล", "cleanOldData: เริ่มทำความสะอาดข้อมูล...");

  // ── 1. กำหนดวันที่ตัด = วันนี้ − 60 วัน ──────────────────
  const today   = new Date();
  const cutoff  = new Date(today);
  cutoff.setDate(cutoff.getDate() - 210);
  const cutoffStr = formatDate(cutoff);

  // ── 2. ดึงข้อมูลทั้งหมด ───────────────────────────────────
  const lastRow = logSheet.getLastRow();
  const numRows = lastRow - 1;
  const data    = logSheet.getRange(2, 1, numRows, 6).getValues();

  // ── 3. หาค่าสูงสุดต่อ Symbol ─────────────────────────────
  // { "VOO": { value: 698.26, rowIndex: 5 } }
  const highestMap = {};  // symbol → { value, rowIdx (0-based ใน data) }

  data.forEach((row, i) => {
    const sym   = String(row[1]).trim();
    const close = parseFloat(row[5]);
    if (!sym || isNaN(close)) return;

    if (!highestMap[sym] || close > highestMap[sym].value) {
      highestMap[sym] = { value: close, rowIdx: i };
    }
  });

  // rowIdx ของแถวที่เป็นค่าสูงสุด (ต้องเก็บไว้ไม่ว่าจะเก่าแค่ไหน)
  const keepRows = new Set(Object.values(highestMap).map(h => h.rowIdx));

  // ── 4. แยกแถวที่ "ลบ" vs "เก็บ" ─────────────────────────
  const rowsToDelete = [];   // เก็บ index (0-based) ที่จะลบ
  const rowsToKeep   = [];   // เก็บไว้ทั้งหมด

  data.forEach((row, i) => {
    const dateVal = row[0];
    const dateStr = dateVal instanceof Date
      ? formatDate(dateVal)
      : String(dateVal).substring(0, 10);

    const isOld     = dateStr < cutoffStr;     // เก่ากว่า 60 วัน
    const isHighest = keepRows.has(i);          // เป็นค่าสูงสุด

    if (isOld && !isHighest) {
      rowsToDelete.push(i);   // ลบ: เก่า และไม่ใช่ค่าสูงสุด
    } else {
      rowsToKeep.push(row);   // เก็บ: ใหม่ หรือเป็นค่าสูงสุด
    }
  });

  if (rowsToDelete.length === 0) {
    const msg = `cleanOldData: ไม่มีข้อมูลที่ต้องลบ (ทั้งหมดอยู่ใน 60 วัน)`;
    writeLog("ข้อมูล", msg);
    Logger.log(msg);
    return;
  }

  // ── 5. เขียนข้อมูลที่เก็บไว้กลับลง Sheet ────────────────
  // ล้าง Sheet แล้วเขียนใหม่ทั้งหมด (เร็วกว่าลบทีละแถว)
  logSheet.clearContents();

  // Header
  logSheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
  logSheet.setFrozenRows(1);

  if (rowsToKeep.length > 0) {
    logSheet.getRange(2, 1, rowsToKeep.length, 6).setValues(rowsToKeep);
  }

  // ── 6. บันทึก Log ─────────────────────────────────────────
  const highestSymbols = Object.entries(highestMap)
    .filter(([sym, h]) => keepRows.has(h.rowIdx) && rowsToDelete.includes(h.rowIdx))
    .map(([sym]) => sym);

  const msg = (
    `cleanOldData: เสร็จแล้ว\n` +
    `ลบ: ${rowsToDelete.length} แถว\n` +
    `เก็บ: ${rowsToKeep.length} แถว\n` +
    `เก็บค่าสูงสุดของ: [${Object.keys(highestMap).join(", ")}]`
  );

  Logger.log(msg);
  writeLog("ข้อมูล", `cleanOldData: ลบ ${rowsToDelete.length} แถว | เก็บ ${rowsToKeep.length} แถว`);
  writeLog("ข้อมูล", `เก็บค่าสูงสุดของ: [${Object.keys(highestMap).join(", ")}]`);

  sendTelegram(
    `🧹 <b>Clean Data เสร็จแล้ว</b>\n` +
    `🗑 ลบ: ${rowsToDelete.length} แถว (เก่ากว่า ${cutoffStr})\n` +
    `✅ เก็บ: ${rowsToKeep.length} แถว\n` +
    `📌 ค่าสูงสุดของทุก Symbol ถูกเก็บไว้`
  );

  // แจ้งถ้ารันด้วยมือ
  try {
    SpreadsheetApp.getUi().alert(
      `🧹 Clean Data เสร็จแล้ว!\n\n` +
      `🗑 ลบออก: ${rowsToDelete.length} แถว\n` +
      `✅ เก็บไว้: ${rowsToKeep.length} แถว\n` +
      `📌 ค่าสูงสุดของ ${Object.keys(highestMap).length} symbols ถูกเก็บไว้ทั้งหมด`
    );
  } catch(e) { /* รันจาก Trigger — ไม่มี UI */ }
}

// ============================================================
//  setupCleanTrigger()
//  ตั้ง Trigger รัน cleanOldData() ทุกต้นเดือน
//  รันครั้งเดียวเพื่อตั้ง Trigger
// ============================================================
function setupCleanTrigger() {
  // ลบ Trigger เก่า
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "cleanOldData")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // ตั้งใหม่ — ทุกวันที่ 1 ของเดือน เวลา 03:00-04:00
  ScriptApp.newTrigger("cleanOldData")
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();

  Logger.log("setupCleanTrigger: ตั้ง Trigger รัน cleanOldData ทุกวันที่ 1 ของเดือน 03:00");

  try {
    SpreadsheetApp.getUi().alert(
      "✅ Trigger ตั้งแล้ว!\n\n" +
      "cleanOldData() จะรันอัตโนมัติทุกวันที่ 1 ของเดือน 03:00\n\n" +
      "หรือรัน cleanOldData() ด้วยมือได้ตลอดเวลา"
    );
  } catch(e) {}
}






## STOCK_PRICE.gs
//  STOCK_PRICE.gs
//  STOCK_PRICE_DATABASE — Google Apps Script (Final Version)
//  อัปเดต: 2026-06-24
// ============================================================

// ──────────────────────────────────────────────────────────────
//  ⚙️ CONFIG — แก้ค่าตรงนี้เท่านั้น
// ──────────────────────────────────────────────────────────────
const CONFIG = {

  // ✅ ไฟล์หลัก   https://docs.google.com/spreadsheets/d/1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo/edit?usp=drive_link
  PORTFOLIO_FILE_ID: "1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo",

  // ✅ Sheet ต้นทาง Ticker + เขียน ATR14 กลับ
  BACKEND_ATR_SHEET:  "Backend_ATR",
  US_TICKER_COL:      1,   // col A = Ticker หุ้นสหรัฐ
  US_ATR_WRITE_COL:   2,   // col B = เขียน ATR14 หุ้นสหรัฐ
  TH_TICKER_COL:      5,   // col E = Ticker หุ้นไทย
  TH_ATR_WRITE_COL:   6,   // col F = เขียน ATR14 หุ้นไทย
  DATA_START_ROW:     2,   // แถวเริ่มข้อมูล (ข้าม Header แถว 1)

  // ──────────────────────────────────────────────────────────
  //  📌 HIGHEST CLOSE SYNC
  //  กำหนดได้หลายปลายทาง — เพิ่ม/ลด object ใน array ได้เลย
  //  fileId: "" = ไฟล์หลัก (PORTFOLIO_FILE_ID) / ใส่ ID อื่นถ้าต้องการ
  //  symbolCol: คอลัมน์ที่มี Symbol ใน sheet ปลายทาง
  //  writeCol:  คอลัมน์ที่ต้องการให้เขียน Highest Close
  //  market:    "US" | "TH" | "ALL"  (กรองเฉพาะ US หรือ TH หรือทั้งหมด)
  // ──────────────────────────────────────────────────────────
  HIGHEST_CLOSE_TARGETS: [
    //  fileId:    "" = ไฟล์หลัก (PORTFOLIO_FILE_ID) / ใส่ ID อื่นถ้าต้องการไฟล์อื่น
    //  sheet:     ชื่อ Sheet ปลายทาง
    //  symbolCol: คอลัมน์ที่มี Symbol ใน sheet นั้น (A=1, B=2, ...)
    //  writeCol:  คอลัมน์ที่ต้องการให้เขียน Highest Close
    //  market:    "US" | "TH" | "ALL"

    { fileId: "1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo", sheet: "📊 ATR_Portfolio US", symbolCol: 1, writeCol: 4, market: "US" },
    //             ไฟล์หลัก  ชื่อ Sheet           Symbol col A   Highest Close col D

    { fileId: "1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo", sheet: "📊 ATR_Portfolio TH", symbolCol: 1, writeCol: 4, market: "TH" },
    //             ไฟล์หลัก  ชื่อ Sheet           Symbol col A   Highest Close col D

    // ⬇️ เพิ่มปลายทางได้อีกได้เลย เช่น
    // { fileId: "", sheet: "Backend_ATR", symbolCol: 1, writeCol: 3, market: "US" },
  ],

  // Telegram (ไม่บังคับ)
  TELEGRAM_BOT_TOKEN: "",
  TELEGRAM_CHAT_ID:   "",

  // ATR14 รายสัปดาห์ — 0=อาทิตย์ 1=จันทร์ ... 6=เสาร์
  ATR_DAY_OF_WEEK: 1,

  // ATR กี่วัน
  PERIOD_DAY_OF_DAY : 14 ,

};

// ──────────────────────────────────────────────────────────────
//  ยูทิลิตี้
// ──────────────────────────────────────────────────────────────

function getDB() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function sendTelegram(msg) {
  if (!CONFIG.TELEGRAM_BOT_TOKEN) return;
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: CONFIG.TELEGRAM_CHAT_ID, text: msg, parse_mode: "HTML" }),
      muteHttpExceptions: true,
    }
  );
}

/** เปิดไฟล์ — "" หรือไม่ใส่ = ไฟล์หลัก PORTFOLIO_FILE_ID */
function openFile(fileId) {
  const id = fileId || CONFIG.PORTFOLIO_FILE_ID;
  return SpreadsheetApp.openById(id);
}

// ──────────────────────────────────────────────────────────────
//  getTickers() — ดึงจาก Backend_ATR ในไฟล์หลักอัตโนมัติ
// ──────────────────────────────────────────────────────────────
function getTickers() {
  const portfolio = openFile();
  const sheet     = portfolio.getSheetByName(CONFIG.BACKEND_ATR_SHEET);
  if (!sheet) throw new Error(`ไม่พบ Sheet: ${CONFIG.BACKEND_ATR_SHEET}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { us: [], th: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const maxCol  = Math.max(CONFIG.US_TICKER_COL, CONFIG.TH_TICKER_COL);
  const allData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, maxCol).getValues();

  const us = [], th = [];
  allData.forEach(row => {
    const u = String(row[CONFIG.US_TICKER_COL - 1] || "").trim();
    const t = String(row[CONFIG.TH_TICKER_COL - 1] || "").trim();
    if (u && u !== "Ticker (หุ้นสหรัฐ)") us.push(u);
    if (t && t !== "Ticker (หุ้นไทย)")   th.push(t);
  });

  Logger.log(`Tickers US=[${us}] TH=[${th}]`);
  return { us, th };
}

// ──────────────────────────────────────────────────────────────
//  ดึงราคา Yahoo Finance
// ──────────────────────────────────────────────────────────────
function fetchUSPrice(symbol) {
  try {
    const url  = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?interval=1d&range=1d`;
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const meta = json.chart.result[0].meta;
    return {
      symbol,
      price: meta.regularMarketPrice,
      high:  meta.regularMarketDayHigh,
      low:   meta.regularMarketDayLow,
      close: meta.regularMarketPrice,
    };
  } catch (e) {
    Logger.log(`fetchUSPrice error [${symbol}]: ${e}`);
    return null;
  }
}

function fetchTHPrice(symbol) {
  const sym  = symbol.endsWith(".BK") ? symbol : symbol + ".BK";
  const data = fetchUSPrice(sym);
  if (data) data.symbol = symbol;
  return data;
}

// ──────────────────────────────────────────────────────────────
//  ⚙️ Script 1 : saveDailyPrice()
// ──────────────────────────────────────────────────────────────
function saveDailyPrice() {
  const db    = getDB();
  const sheet = getOrCreateSheet(db, "Daily_Close_Log");
  const today = formatDate(new Date());

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
    sheet.setFrozenRows(1);
  }

  const { us, th } = getTickers();
  const rows = [];

  us.forEach(sym => {
    const d = fetchUSPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  th.forEach(sym => {
    const d = fetchTHPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }

  Logger.log(`saveDailyPrice: ${rows.length} รายการ วันที่ ${today}`);
  sendTelegram(`✅ <b>Price Saved</b>\n📅 ${today}\n📊 ${rows.length} ตัว (US:${us.length} TH:${th.length})`);
}



// ──────────────────────────────────────────────────────────────
//  ⚙️ Script 2 : updateHighestClose()
// ──────────────────────────────────────────────────────────────
//  🔧 PATCH 1 : updateHighestClose()
//  เพิ่ม col C = วันที่ที่เกิด Highest Close
// ──────────────────────────────────────────────────────────────

/**
 * ====================================================
 * 


function updateHighestClose() {
  const db       = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  if (!logSheet || logSheet.getLastRow() < 2) return;

  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();

  // หา Highest Close + วันที่ที่เกิด per Symbol
  const highest = {};  // { sym: { value, date } }

  data.forEach(row => {
    const dateStr = row[0];          // col A = Date
    const sym     = String(row[1]).trim();
    const close   = parseFloat(row[5]);
    if (!sym || isNaN(close)) return;

    if (!highest[sym] || close > highest[sym].value) {
      highest[sym] = { value: close, date: dateStr };
    }
  });

  // เขียนลง Highest_Close_Summary (3 คอลัมน์)
  const sumSheet = getOrCreateSheet(db, "Highest_Close_Summary");
  sumSheet.clearContents();
  sumSheet.appendRow(["Symbol", "Highest Close", "Date"]);
  sumSheet.setFrozenRows(1);

  // จัด style header
  const hRow = sumSheet.getRange(1, 1, 1, 3);
  hRow.setFontWeight("bold");

  Object.entries(highest)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([sym, { value, date }]) => {
      sumSheet.appendRow([sym, value, date]);
    });

  // format col C เป็น date
  if (sumSheet.getLastRow() > 1) {
    sumSheet.getRange(2, 3, sumSheet.getLastRow() - 1, 1)
            .setNumberFormat("yyyy-mm-dd");
  }

  Logger.log(`updateHighestClose: ${Object.keys(highest).length} symbols (พร้อม Date)`);
}

 * ====================================================
 * 
 */


// ──────────────────────────────────────────────────────────────
//  🔧 PATCH 2 : writeLog()  — helper เขียน Trigger_Log
// ──────────────────────────────────────────────────────────────
function writeLog(type, detail) {
  const db    = getDB();
  const sheet = getOrCreateSheet(db, "Trigger_Log");

  // สร้าง Header ถ้ายังไม่มี
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Time", "Type", "Detail"]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 700);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }

  const now  = new Date();
  const tz   = Session.getScriptTimeZone();
  const date = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const time = Utilities.formatDate(now, tz, "HH:mm:ss");

  sheet.appendRow([date, time, type, detail]);
}

// 


// ──────────────────────────────────────────────────────────────
//  ⚙️ Script 3 : syncPortfolio()
//  เขียน Highest Close ไปทุก target ที่กำหนดใน CONFIG
// ──────────────────────────────────────────────────────────────
// ──────────────────────────────────────────────────────────────
//  อัปเดต syncPortfolio() ให้บันทึก log ด้วย
// ──────────────────────────────────────────────────────────────
function syncPortfolio() {
  if (!CONFIG.HIGHEST_CLOSE_TARGETS || CONFIG.HIGHEST_CLOSE_TARGETS.length === 0) {
    writeLog("ข้อมูล", "syncPortfolio: ยังไม่มี HIGHEST_CLOSE_TARGETS — ข้าม");
    return;
  }

  const db       = getDB();
  const sumSheet = db.getSheetByName("Highest_Close_Summary");
  if (!sumSheet || sumSheet.getLastRow() < 2) return;

  const hcMap = {};
  sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, 2).getValues()
    .forEach(r => { hcMap[String(r[0]).trim()] = r[1]; });

  const { us, th } = getTickers();
  const usSet = new Set(us);
  const thSet = new Set(th);

  CONFIG.HIGHEST_CLOSE_TARGETS.forEach(target => {
    try {
      const file  = openFile(target.fileId);
      const sheet = file.getSheetByName(target.sheet);
      if (!sheet) {
        writeLog("⚠️ WARNING", `syncPortfolio: ไม่พบ Sheet "${target.sheet}"`);
        return;
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < CONFIG.DATA_START_ROW) return;

      const numRows  = lastRow - CONFIG.DATA_START_ROW + 1;
      const syms     = sheet.getRange(CONFIG.DATA_START_ROW, target.symbolCol, numRows, 1).getValues();
      const writeData = [];

      syms.forEach((r, i) => {
        const sym = String(r[0]).trim();
        if (!sym) return;
        if (target.market === "US" && !usSet.has(sym)) return;
        if (target.market === "TH" && !thSet.has(sym)) return;
        if (hcMap[sym] !== undefined) {
          writeData.push({ row: CONFIG.DATA_START_ROW + i, val: hcMap[sym] });
        }
      });

      writeData.forEach(({ row, val }) => {
        sheet.getRange(row, target.writeCol).setValue(val);
      });

      const msg = `syncPortfolio: "${target.sheet}" col ${target.writeCol} — เขียน ${writeData.length} แถว`;
      Logger.log(msg);
      writeLog("ข้อมูล", msg);

    } catch (e) {
      writeLog("❌ ERROR", `syncPortfolio [${target.sheet}]: ${e}`);
    }
  });

  sendTelegram(`📤 <b>Highest Close Synced</b>\nเขียนไป ${CONFIG.HIGHEST_CLOSE_TARGETS.length} ปลายทางแล้ว`);
}

// ──────────────────────────────────────────────────────────────
//  อัปเดต calculateATR14() ให้บันทึก log ด้วย
// ──────────────────────────────────────────────────────────────
function calculateATR14() {
  const db       = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  if (!logSheet || logSheet.getLastRow() < 2) {
    writeLog("⚠️ WARNING", "calculateATR14: ไม่มีข้อมูลใน Daily_Close_Log");
    return;
  }

  const data     = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();
  const bySymbol = {};
  data.forEach(row => {
    const sym   = String(row[1]).trim();
    const high  = parseFloat(row[3]);
    const low   = parseFloat(row[4]);
    const close = parseFloat(row[5]);
    if (!sym || isNaN(high) || isNaN(low) || isNaN(close)) return;
    if (!bySymbol[sym]) bySymbol[sym] = [];
    bySymbol[sym].push({ high, low, close });
  });

  const atrSheet = getOrCreateSheet(db, "ATR_History");
  if (atrSheet.getLastRow() === 0) {
    atrSheet.appendRow(["Date", "Symbol", "ATR14"]);
    atrSheet.setFrozenRows(1);
  }

  const dateStr = formatDate(new Date());
  const newRows = [];
  const atrMap  = {};
  const period  = CONFIG.PERIOD_DAY_OF_DAY;

  Object.entries(bySymbol).forEach(([sym, bars]) => {
    if (bars.length < period + 1) return;
    const tr = [];
    for (let i = 1; i < bars.length; i++) {
      const c = bars[i], p = bars[i - 1];
      tr.push(Math.max(c.high - c.low, Math.abs(c.high - p.close), Math.abs(c.low - p.close)));
    }
    if (tr.length < period) return;
    let atr = tr.slice(0, period).reduce((a, b) => a + b, 0) / period;
    for (let i = period; i < tr.length; i++) atr = (atr * (period - 1) + tr[i]) / period;
    atrMap[sym] = Math.round(atr * 100) / 100;
    newRows.push([dateStr, sym, atrMap[sym]]);
  });

  if (newRows.length > 0) {
    atrSheet.getRange(atrSheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  }

  // เขียนกลับ Backend_ATR
  const portfolio = openFile();
  const sheet     = portfolio.getSheetByName(CONFIG.BACKEND_ATR_SHEET);
  if (sheet && sheet.getLastRow() >= CONFIG.DATA_START_ROW) {
    const numRows = sheet.getLastRow() - CONFIG.DATA_START_ROW + 1;
    const maxCol  = Math.max(CONFIG.US_TICKER_COL, CONFIG.TH_TICKER_COL);
    const allData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, maxCol).getValues();
    allData.forEach((row, i) => {
      const rowNum = CONFIG.DATA_START_ROW + i;
      const usSym  = String(row[CONFIG.US_TICKER_COL - 1] || "").trim();
      const thSym  = String(row[CONFIG.TH_TICKER_COL - 1] || "").trim();
      if (usSym && atrMap[usSym] !== undefined) sheet.getRange(rowNum, CONFIG.US_ATR_WRITE_COL).setValue(atrMap[usSym]);
      if (thSym && atrMap[thSym] !== undefined) sheet.getRange(rowNum, CONFIG.TH_ATR_WRITE_COL).setValue(atrMap[thSym]);
    });
  }

  const msg = `calculateATR14: ${newRows.length} symbols → ATR_History + Backend_ATR (ค่าจริง)`;
  Logger.log(msg);
  writeLog("ข้อมูล", msg);
  sendTelegram(`📐 <b>ATR14 Updated</b>\n📅 ${dateStr}\n✅ ${newRows.length} ตัว`);
}

// ──────────────────────────────────────────────────────────────
//  Master Runner
// ──────────────────────────────────────────────────────────────
// ──────────────────────────────────────────────────────────────
//  🔧 PATCH 3 : runDailyJob()  — บันทึก Trigger_Log ทุกขั้นตอน
// ──────────────────────────────────────────────────────────────
function runDailyJob() {
  writeLog("ประกาศ", "เริ่มการดำเนินการแล้ว");

  try {
    writeLog("ข้อมูล", "=== START runDailyJob ===");

   // ── Step 0 : backfillNewTickers ── 
   backfillNewTickers()


    // ── Step 1 : saveDailyPrice ──
    writeLog("ข้อมูล", "เริ่ม saveDailyPrice()...");
    saveDailyPrice();

    // ── Step 2 : updateHighestClose ──
    writeLog("ข้อมูล", "เริ่ม updateHighestClose()...");
    updateHighestClose();

    // ── Step 3 : syncPortfolio ──
    writeLog("ข้อมูล", "เริ่ม syncPortfolio()...");
    syncPortfolio();

    // ── Step 4 : calculateATR14 (เฉพาะวันจันทร์ คืนวันอาทิตย์) ──
    const day = new Date().getDay();
    if (day === CONFIG.ATR_DAY_OF_WEEK) {
      writeLog("ข้อมูล", "เริ่ม calculateATR14() (วันอาทิตย์)...");
      calculateATR14();
    } else {
      writeLog("ข้อมูล", `ข้าม calculateATR14() — วันนี้ไม่ใช่วันอาทิตย์ (day=${day})`);
    }

    writeLog("ข้อมูล", "=== END runDailyJob ===");
    writeLog("ประกาศ", "ดำเนินการเสร็จแล้ว ✅");

  } catch (e) {
    writeLog("❌ ERROR", String(e));
    Logger.log("runDailyJob ERROR: " + e);
    sendTelegram(`❌ <b>Error: runDailyJob</b>\n${e}`);
  }
}

// ──────────────────────────────────────────────────────────────
//  อัปเดต saveDailyPrice() ให้บันทึก log ด้วย
// ──────────────────────────────────────────────────────────────
function saveDailyPrice() {
  const db    = getDB();
  const sheet = getOrCreateSheet(db, "Daily_Close_Log");
  const today = formatDate(new Date());

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
    sheet.setFrozenRows(1);
  }

  const { us, th } = getTickers();
  writeLog("ข้อมูล", `Tickers US=[${us.join(",")}] TH=[${th.join(",")}]`);

  const rows = [];

  us.forEach(sym => {
    const d = fetchUSPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  th.forEach(sym => {
    const d = fetchTHPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }

  const msg = `saveDailyPrice: ${rows.length} รายการ วันที่ ${today}`;
  Logger.log(msg);
  writeLog("ข้อมูล", msg);
  sendTelegram(`✅ <b>Price Saved</b>\n📅 ${today}\n📊 ${rows.length} ตัว (US:${us.length} TH:${th.length})`);
}


// ──────────────────────────────────────────────────────────────
//  🛠️ Setup — รันครั้งเดียวตอนติดตั้ง
// ──────────────────────────────────────────────────────────────
function setupDatabase() {
  const db = getDB();

  [
    { name: "Daily_Close_Log",       header: ["Date","Symbol","Price","High","Low","Close"] },
    { name: "Highest_Close_Summary", header: ["Symbol","Highest Close"] },
    { name: "ATR_History",           header: ["Date","Symbol","ATR14"] },
  ].forEach(({ name, header }) => {
    const s = getOrCreateSheet(db, name);
    if (s.getLastRow() === 0) { s.appendRow(header); s.setFrozenRows(1); }
  });

  // ทดสอบการเชื่อมต่อ
  let tickerMsg = "";
  try {
    const { us, th } = getTickers();
    tickerMsg = `\n\n🇺🇸 US (${us.length}): ${us.join(", ")}\n🇹🇭 TH (${th.length}): ${th.join(", ")}`;
  } catch (e) {
    SpreadsheetApp.getUi().alert("⚠️ เชื่อมต่อไฟล์หลักไม่ได้:\n" + e);
    return;
  }

  // Trigger
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runDailyJob")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("runDailyJob").timeBased().everyDays(1).atHour(2).create();

  SpreadsheetApp.getUi().alert(
    "✅ Setup เสร็จแล้ว!\n" +
    "📋 Sheets: Daily_Close_Log, Highest_Close_Summary, ATR_History\n" +
    "⏰ Trigger: ทุกวัน 02:00 (runDailyJob)" +
    tickerMsg
  );
}


// ============================================================
//  STOCK_PRICE.gs
//  STOCK_PRICE_DATABASE — Google Apps Script (Final Version)
//  อัปเดต: 2026-06-24
// ============================================================

// ──────────────────────────────────────────────────────────────
//  ⚙️ CONFIG — แก้ค่าตรงนี้เท่านั้น
// ──────────────────────────────────────────────────────────────
const CONFIG = {

  // ✅ ไฟล์หลัก   https://docs.google.com/spreadsheets/d/1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo/edit?usp=drive_link
  PORTFOLIO_FILE_ID: "1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo",

  // ✅ Sheet ต้นทาง Ticker + เขียน ATR14 กลับ
  BACKEND_ATR_SHEET:  "Backend_ATR",
  US_TICKER_COL:      1,   // col A = Ticker หุ้นสหรัฐ
  US_ATR_WRITE_COL:   2,   // col B = เขียน ATR14 หุ้นสหรัฐ
  TH_TICKER_COL:      5,   // col E = Ticker หุ้นไทย
  TH_ATR_WRITE_COL:   6,   // col F = เขียน ATR14 หุ้นไทย
  DATA_START_ROW:     2,   // แถวเริ่มข้อมูล (ข้าม Header แถว 1)

  // ──────────────────────────────────────────────────────────
  //  📌 HIGHEST CLOSE SYNC
  //  กำหนดได้หลายปลายทาง — เพิ่ม/ลด object ใน array ได้เลย
  //  fileId: "" = ไฟล์หลัก (PORTFOLIO_FILE_ID) / ใส่ ID อื่นถ้าต้องการ
  //  symbolCol: คอลัมน์ที่มี Symbol ใน sheet ปลายทาง
  //  writeCol:  คอลัมน์ที่ต้องการให้เขียน Highest Close
  //  market:    "US" | "TH" | "ALL"  (กรองเฉพาะ US หรือ TH หรือทั้งหมด)
  // ──────────────────────────────────────────────────────────
  HIGHEST_CLOSE_TARGETS: [
    //  fileId:    "" = ไฟล์หลัก (PORTFOLIO_FILE_ID) / ใส่ ID อื่นถ้าต้องการไฟล์อื่น
    //  sheet:     ชื่อ Sheet ปลายทาง
    //  symbolCol: คอลัมน์ที่มี Symbol ใน sheet นั้น (A=1, B=2, ...)
    //  writeCol:  คอลัมน์ที่ต้องการให้เขียน Highest Close
    //  market:    "US" | "TH" | "ALL"

    { fileId: "1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo", sheet: "📊 ATR_Portfolio US", symbolCol: 1, writeCol: 4, market: "US" },
    //             ไฟล์หลัก  ชื่อ Sheet           Symbol col A   Highest Close col D

    { fileId: "1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo", sheet: "📊 ATR_Portfolio TH", symbolCol: 1, writeCol: 4, market: "TH" },
    //             ไฟล์หลัก  ชื่อ Sheet           Symbol col A   Highest Close col D

    // ⬇️ เพิ่มปลายทางได้อีกได้เลย เช่น
    // { fileId: "", sheet: "Backend_ATR", symbolCol: 1, writeCol: 3, market: "US" },
  ],

  // Telegram (ไม่บังคับ)
  TELEGRAM_BOT_TOKEN: "",
  TELEGRAM_CHAT_ID:   "",

  // ATR14 รายสัปดาห์ — 0=อาทิตย์ 1=จันทร์ ... 6=เสาร์
  ATR_DAY_OF_WEEK: 1,

  // ATR กี่วัน
  PERIOD_DAY_OF_DAY : 14 ,

};

// ──────────────────────────────────────────────────────────────
//  ยูทิลิตี้
// ──────────────────────────────────────────────────────────────

function getDB() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function sendTelegram(msg) {
  if (!CONFIG.TELEGRAM_BOT_TOKEN) return;
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: CONFIG.TELEGRAM_CHAT_ID, text: msg, parse_mode: "HTML" }),
      muteHttpExceptions: true,
    }
  );
}

/** เปิดไฟล์ — "" หรือไม่ใส่ = ไฟล์หลัก PORTFOLIO_FILE_ID */
function openFile(fileId) {
  const id = fileId || CONFIG.PORTFOLIO_FILE_ID;
  return SpreadsheetApp.openById(id);
}

// ──────────────────────────────────────────────────────────────
//  getTickers() — ดึงจาก Backend_ATR ในไฟล์หลักอัตโนมัติ
// ──────────────────────────────────────────────────────────────
function getTickers() {
  const portfolio = openFile();
  const sheet     = portfolio.getSheetByName(CONFIG.BACKEND_ATR_SHEET);
  if (!sheet) throw new Error(`ไม่พบ Sheet: ${CONFIG.BACKEND_ATR_SHEET}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { us: [], th: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const maxCol  = Math.max(CONFIG.US_TICKER_COL, CONFIG.TH_TICKER_COL);
  const allData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, maxCol).getValues();

  const us = [], th = [];
  allData.forEach(row => {
    const u = String(row[CONFIG.US_TICKER_COL - 1] || "").trim();
    const t = String(row[CONFIG.TH_TICKER_COL - 1] || "").trim();
    if (u && u !== "Ticker (หุ้นสหรัฐ)") us.push(u);
    if (t && t !== "Ticker (หุ้นไทย)")   th.push(t);
  });

  Logger.log(`Tickers US=[${us}] TH=[${th}]`);
  return { us, th };
}

// ──────────────────────────────────────────────────────────────
//  ดึงราคา Yahoo Finance
// ──────────────────────────────────────────────────────────────
function fetchUSPrice(symbol) {
  try {
    const url  = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?interval=1d&range=1d`;
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const meta = json.chart.result[0].meta;
    return {
      symbol,
      price: meta.regularMarketPrice,
      high:  meta.regularMarketDayHigh,
      low:   meta.regularMarketDayLow,
      close: meta.regularMarketPrice,
    };
  } catch (e) {
    Logger.log(`fetchUSPrice error [${symbol}]: ${e}`);
    return null;
  }
}

function fetchTHPrice(symbol) {
  const sym  = symbol.endsWith(".BK") ? symbol : symbol + ".BK";
  const data = fetchUSPrice(sym);
  if (data) data.symbol = symbol;
  return data;
}

// ──────────────────────────────────────────────────────────────
//  ⚙️ Script 1 : saveDailyPrice()
// ──────────────────────────────────────────────────────────────
function saveDailyPrice() {
  const db    = getDB();
  const sheet = getOrCreateSheet(db, "Daily_Close_Log");
  const today = formatDate(new Date());

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
    sheet.setFrozenRows(1);
  }

  const { us, th } = getTickers();
  const rows = [];

  us.forEach(sym => {
    const d = fetchUSPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  th.forEach(sym => {
    const d = fetchTHPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }

  Logger.log(`saveDailyPrice: ${rows.length} รายการ วันที่ ${today}`);
  sendTelegram(`✅ <b>Price Saved</b>\n📅 ${today}\n📊 ${rows.length} ตัว (US:${us.length} TH:${th.length})`);
}



// ──────────────────────────────────────────────────────────────
//  ⚙️ Script 2 : updateHighestClose()
// ──────────────────────────────────────────────────────────────
//  🔧 PATCH 1 : updateHighestClose()
//  เพิ่ม col C = วันที่ที่เกิด Highest Close
// ──────────────────────────────────────────────────────────────

/**
 * ====================================================
 * 


function updateHighestClose() {
  const db       = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  if (!logSheet || logSheet.getLastRow() < 2) return;

  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();

  // หา Highest Close + วันที่ที่เกิด per Symbol
  const highest = {};  // { sym: { value, date } }

  data.forEach(row => {
    const dateStr = row[0];          // col A = Date
    const sym     = String(row[1]).trim();
    const close   = parseFloat(row[5]);
    if (!sym || isNaN(close)) return;

    if (!highest[sym] || close > highest[sym].value) {
      highest[sym] = { value: close, date: dateStr };
    }
  });

  // เขียนลง Highest_Close_Summary (3 คอลัมน์)
  const sumSheet = getOrCreateSheet(db, "Highest_Close_Summary");
  sumSheet.clearContents();
  sumSheet.appendRow(["Symbol", "Highest Close", "Date"]);
  sumSheet.setFrozenRows(1);

  // จัด style header
  const hRow = sumSheet.getRange(1, 1, 1, 3);
  hRow.setFontWeight("bold");

  Object.entries(highest)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([sym, { value, date }]) => {
      sumSheet.appendRow([sym, value, date]);
    });

  // format col C เป็น date
  if (sumSheet.getLastRow() > 1) {
    sumSheet.getRange(2, 3, sumSheet.getLastRow() - 1, 1)
            .setNumberFormat("yyyy-mm-dd");
  }

  Logger.log(`updateHighestClose: ${Object.keys(highest).length} symbols (พร้อม Date)`);
}

 * ====================================================
 * 
 */


// ──────────────────────────────────────────────────────────────
//  🔧 PATCH 2 : writeLog()  — helper เขียน Trigger_Log
// ──────────────────────────────────────────────────────────────
function writeLog(type, detail) {
  const db    = getDB();
  const sheet = getOrCreateSheet(db, "Trigger_Log");

  // สร้าง Header ถ้ายังไม่มี
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Time", "Type", "Detail"]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 700);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }

  const now  = new Date();
  const tz   = Session.getScriptTimeZone();
  const date = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const time = Utilities.formatDate(now, tz, "HH:mm:ss");

  sheet.appendRow([date, time, type, detail]);
}

// 


// ──────────────────────────────────────────────────────────────
//  ⚙️ Script 3 : syncPortfolio()
//  เขียน Highest Close ไปทุก target ที่กำหนดใน CONFIG
// ──────────────────────────────────────────────────────────────
// ──────────────────────────────────────────────────────────────
//  อัปเดต syncPortfolio() ให้บันทึก log ด้วย
// ──────────────────────────────────────────────────────────────
function syncPortfolio() {
  if (!CONFIG.HIGHEST_CLOSE_TARGETS || CONFIG.HIGHEST_CLOSE_TARGETS.length === 0) {
    writeLog("ข้อมูล", "syncPortfolio: ยังไม่มี HIGHEST_CLOSE_TARGETS — ข้าม");
    return;
  }

  const db       = getDB();
  const sumSheet = db.getSheetByName("Highest_Close_Summary");
  if (!sumSheet || sumSheet.getLastRow() < 2) return;

  const hcMap = {};
  sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, 2).getValues()
    .forEach(r => { hcMap[String(r[0]).trim()] = r[1]; });

  const { us, th } = getTickers();
  const usSet = new Set(us);
  const thSet = new Set(th);

  CONFIG.HIGHEST_CLOSE_TARGETS.forEach(target => {
    try {
      const file  = openFile(target.fileId);
      const sheet = file.getSheetByName(target.sheet);
      if (!sheet) {
        writeLog("⚠️ WARNING", `syncPortfolio: ไม่พบ Sheet "${target.sheet}"`);
        return;
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < CONFIG.DATA_START_ROW) return;

      const numRows  = lastRow - CONFIG.DATA_START_ROW + 1;
      const syms     = sheet.getRange(CONFIG.DATA_START_ROW, target.symbolCol, numRows, 1).getValues();
      const writeData = [];

      syms.forEach((r, i) => {
        const sym = String(r[0]).trim();
        if (!sym) return;
        if (target.market === "US" && !usSet.has(sym)) return;
        if (target.market === "TH" && !thSet.has(sym)) return;
        if (hcMap[sym] !== undefined) {
          writeData.push({ row: CONFIG.DATA_START_ROW + i, val: hcMap[sym] });
        }
      });

      writeData.forEach(({ row, val }) => {
        sheet.getRange(row, target.writeCol).setValue(val);
      });

      const msg = `syncPortfolio: "${target.sheet}" col ${target.writeCol} — เขียน ${writeData.length} แถว`;
      Logger.log(msg);
      writeLog("ข้อมูล", msg);

    } catch (e) {
      writeLog("❌ ERROR", `syncPortfolio [${target.sheet}]: ${e}`);
    }
  });

  sendTelegram(`📤 <b>Highest Close Synced</b>\nเขียนไป ${CONFIG.HIGHEST_CLOSE_TARGETS.length} ปลายทางแล้ว`);
}

// ──────────────────────────────────────────────────────────────
//  อัปเดต calculateATR14() ให้บันทึก log ด้วย
// ──────────────────────────────────────────────────────────────
function calculateATR14() {
  const db       = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  if (!logSheet || logSheet.getLastRow() < 2) {
    writeLog("⚠️ WARNING", "calculateATR14: ไม่มีข้อมูลใน Daily_Close_Log");
    return;
  }

  const data     = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();
  const bySymbol = {};
  data.forEach(row => {
    const sym   = String(row[1]).trim();
    const high  = parseFloat(row[3]);
    const low   = parseFloat(row[4]);
    const close = parseFloat(row[5]);
    if (!sym || isNaN(high) || isNaN(low) || isNaN(close)) return;
    if (!bySymbol[sym]) bySymbol[sym] = [];
    bySymbol[sym].push({ high, low, close });
  });

  const atrSheet = getOrCreateSheet(db, "ATR_History");
  if (atrSheet.getLastRow() === 0) {
    atrSheet.appendRow(["Date", "Symbol", "ATR14"]);
    atrSheet.setFrozenRows(1);
  }

  const dateStr = formatDate(new Date());
  const newRows = [];
  const atrMap  = {};
  const period  = CONFIG.PERIOD_DAY_OF_DAY;

  Object.entries(bySymbol).forEach(([sym, bars]) => {
    if (bars.length < period + 1) return;
    const tr = [];
    for (let i = 1; i < bars.length; i++) {
      const c = bars[i], p = bars[i - 1];
      tr.push(Math.max(c.high - c.low, Math.abs(c.high - p.close), Math.abs(c.low - p.close)));
    }
    if (tr.length < period) return;
    let atr = tr.slice(0, period).reduce((a, b) => a + b, 0) / period;
    for (let i = period; i < tr.length; i++) atr = (atr * (period - 1) + tr[i]) / period;
    atrMap[sym] = Math.round(atr * 100) / 100;
    newRows.push([dateStr, sym, atrMap[sym]]);
  });

  if (newRows.length > 0) {
    atrSheet.getRange(atrSheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  }

  // เขียนกลับ Backend_ATR
  const portfolio = openFile();
  const sheet     = portfolio.getSheetByName(CONFIG.BACKEND_ATR_SHEET);
  if (sheet && sheet.getLastRow() >= CONFIG.DATA_START_ROW) {
    const numRows = sheet.getLastRow() - CONFIG.DATA_START_ROW + 1;
    const maxCol  = Math.max(CONFIG.US_TICKER_COL, CONFIG.TH_TICKER_COL);
    const allData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, maxCol).getValues();
    allData.forEach((row, i) => {
      const rowNum = CONFIG.DATA_START_ROW + i;
      const usSym  = String(row[CONFIG.US_TICKER_COL - 1] || "").trim();
      const thSym  = String(row[CONFIG.TH_TICKER_COL - 1] || "").trim();
      if (usSym && atrMap[usSym] !== undefined) sheet.getRange(rowNum, CONFIG.US_ATR_WRITE_COL).setValue(atrMap[usSym]);
      if (thSym && atrMap[thSym] !== undefined) sheet.getRange(rowNum, CONFIG.TH_ATR_WRITE_COL).setValue(atrMap[thSym]);
    });
  }

  const msg = `calculateATR14: ${newRows.length} symbols → ATR_History + Backend_ATR (ค่าจริง)`;
  Logger.log(msg);
  writeLog("ข้อมูล", msg);
  sendTelegram(`📐 <b>ATR14 Updated</b>\n📅 ${dateStr}\n✅ ${newRows.length} ตัว`);
}

// ──────────────────────────────────────────────────────────────
//  Master Runner
// ──────────────────────────────────────────────────────────────
// ──────────────────────────────────────────────────────────────
//  🔧 PATCH 3 : runDailyJob()  — บันทึก Trigger_Log ทุกขั้นตอน
// ──────────────────────────────────────────────────────────────
function runDailyJob() {
  writeLog("ประกาศ", "เริ่มการดำเนินการแล้ว");

  try {
    writeLog("ข้อมูล", "=== START runDailyJob ===");

   // ── Step 0 : backfillNewTickers ── 
   backfillNewTickers()


    // ── Step 1 : saveDailyPrice ──
    writeLog("ข้อมูล", "เริ่ม saveDailyPrice()...");
    saveDailyPrice();

    // ── Step 2 : updateHighestClose ──
    writeLog("ข้อมูล", "เริ่ม updateHighestClose()...");
    updateHighestClose();

    // ── Step 3 : syncPortfolio ──
    writeLog("ข้อมูล", "เริ่ม syncPortfolio()...");
    syncPortfolio();

    // ── Step 4 : calculateATR14 (เฉพาะวันจันทร์ คืนวันอาทิตย์) ──
    const day = new Date().getDay();
    if (day === CONFIG.ATR_DAY_OF_WEEK) {
      writeLog("ข้อมูล", "เริ่ม calculateATR14() (วันอาทิตย์)...");
      calculateATR14();
    } else {
      writeLog("ข้อมูล", `ข้าม calculateATR14() — วันนี้ไม่ใช่วันอาทิตย์ (day=${day})`);
    }

    writeLog("ข้อมูล", "=== END runDailyJob ===");
    writeLog("ประกาศ", "ดำเนินการเสร็จแล้ว ✅");

  } catch (e) {
    writeLog("❌ ERROR", String(e));
    Logger.log("runDailyJob ERROR: " + e);
    sendTelegram(`❌ <b>Error: runDailyJob</b>\n${e}`);
  }
}

// ──────────────────────────────────────────────────────────────
//  อัปเดต saveDailyPrice() ให้บันทึก log ด้วย
// ──────────────────────────────────────────────────────────────
function saveDailyPrice() {
  const db    = getDB();
  const sheet = getOrCreateSheet(db, "Daily_Close_Log");
  const today = formatDate(new Date());

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
    sheet.setFrozenRows(1);
  }

  const { us, th } = getTickers();
  writeLog("ข้อมูล", `Tickers US=[${us.join(",")}] TH=[${th.join(",")}]`);

  const rows = [];

  us.forEach(sym => {
    const d = fetchUSPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  th.forEach(sym => {
    const d = fetchTHPrice(sym);
    if (d) rows.push([today, d.symbol, d.price, d.high, d.low, d.close]);
    Utilities.sleep(300);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }

  const msg = `saveDailyPrice: ${rows.length} รายการ วันที่ ${today}`;
  Logger.log(msg);
  writeLog("ข้อมูล", msg);
  sendTelegram(`✅ <b>Price Saved</b>\n📅 ${today}\n📊 ${rows.length} ตัว (US:${us.length} TH:${th.length})`);
}


// ──────────────────────────────────────────────────────────────
//  🛠️ Setup — รันครั้งเดียวตอนติดตั้ง
// ──────────────────────────────────────────────────────────────
function setupDatabase() {
  const db = getDB();

  [
    { name: "Daily_Close_Log",       header: ["Date","Symbol","Price","High","Low","Close"] },
    { name: "Highest_Close_Summary", header: ["Symbol","Highest Close"] },
    { name: "ATR_History",           header: ["Date","Symbol","ATR14"] },
  ].forEach(({ name, header }) => {
    const s = getOrCreateSheet(db, name);
    if (s.getLastRow() === 0) { s.appendRow(header); s.setFrozenRows(1); }
  });

  // ทดสอบการเชื่อมต่อ
  let tickerMsg = "";
  try {
    const { us, th } = getTickers();
    tickerMsg = `\n\n🇺🇸 US (${us.length}): ${us.join(", ")}\n🇹🇭 TH (${th.length}): ${th.join(", ")}`;
  } catch (e) {
    SpreadsheetApp.getUi().alert("⚠️ เชื่อมต่อไฟล์หลักไม่ได้:\n" + e);
    return;
  }

  // Trigger
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runDailyJob")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("runDailyJob").timeBased().everyDays(1).atHour(2).create();

  SpreadsheetApp.getUi().alert(
    "✅ Setup เสร็จแล้ว!\n" +
    "📋 Sheets: Daily_Close_Log, Highest_Close_Summary, ATR_History\n" +
    "⏰ Trigger: ทุกวัน 02:00 (runDailyJob)" +
    tickerMsg
  );
}

## Setup Guide.html
<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>STOCK_PRICE_DATABASE — Setup Guide</title>
<style>
  :root {
    --bg: #0d1117;
    --surface: #161b22;
    --surface2: #1c2128;
    --border: #21262d;
    --accent: #58a6ff;
    --green: #3fb950;
    --yellow: #e3b341;
    --red: #f85149;
    --purple: #bc8cff;
    --muted: #8b949e;
    --text: #e6edf3;
    --mono: 'JetBrains Mono', 'Fira Code', monospace;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    font-size: 15px;
    line-height: 1.75;
    padding: 40px 20px 100px;
    max-width: 820px;
    margin: 0 auto;
  }

  /* ── Header ── */
  .hero {
    border-bottom: 1px solid var(--border);
    padding-bottom: 28px;
    margin-bottom: 40px;
  }
  .hero-label {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 10px;
  }
  .hero h1 {
    font-size: 28px;
    font-weight: 700;
    color: var(--text);
    letter-spacing: -0.5px;
    margin-bottom: 8px;
  }
  .hero p { color: var(--muted); font-size: 14px; }

  /* ── Section titles ── */
  h2 {
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--muted);
    margin: 44px 0 16px;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  h2::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border);
  }

  /* ── Flow diagram ── */
  .flow {
    display: flex;
    align-items: stretch;
    gap: 0;
    margin: 0 0 32px;
    overflow-x: auto;
  }
  .flow-node {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 18px;
    min-width: 140px;
    flex-shrink: 0;
  }
  .flow-node .fn { font-size: 11px; color: var(--muted); margin-bottom: 2px; }
  .flow-node .label { font-size: 13px; font-weight: 600; color: var(--text); }
  .flow-node .sub { font-size: 11px; color: var(--muted); margin-top: 3px; }
  .flow-node.accent { border-color: var(--accent); }
  .flow-node.green  { border-color: var(--green); }
  .flow-node.yellow { border-color: var(--yellow); }
  .flow-arrow {
    display: flex;
    align-items: center;
    padding: 0 6px;
    color: var(--muted);
    font-size: 18px;
    flex-shrink: 0;
  }

  /* ── Sheets ── */
  .sheet-grid { display: grid; gap: 10px; margin-bottom: 8px; }
  .sheet {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px 20px;
    display: grid;
    grid-template-columns: 180px 1fr;
    gap: 0 20px;
    align-items: start;
  }
  .sheet-name {
    font-family: var(--mono);
    font-size: 13px;
    font-weight: 700;
    color: var(--green);
    padding-top: 1px;
  }
  .sheet-info { }
  .sheet-desc { font-size: 13.5px; color: var(--muted); margin-bottom: 6px; }
  .cols { display: flex; flex-wrap: wrap; gap: 6px; }
  .col-tag {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 5px;
    padding: 2px 8px;
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--accent);
  }

  /* ── Steps ── */
  .steps { display: grid; gap: 12px; }
  .step {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 20px 22px;
    display: grid;
    grid-template-columns: 32px 1fr;
    gap: 0 16px;
  }
  .step-num {
    width: 32px;
    height: 32px;
    background: var(--accent);
    color: #000;
    font-size: 13px;
    font-weight: 800;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-top: 1px;
    flex-shrink: 0;
  }
  .step-body { }
  .step-title {
    font-size: 15px;
    font-weight: 700;
    color: var(--text);
    margin-bottom: 8px;
  }
  .step p, .step li { font-size: 13.5px; color: var(--muted); }
  .step ul { margin: 6px 0 0 18px; }
  .step li { margin-bottom: 3px; }

  /* ── Code ── */
  code {
    background: #1f2937;
    color: #79c0ff;
    font-family: var(--mono);
    font-size: 12.5px;
    padding: 2px 7px;
    border-radius: 4px;
  }
  pre {
    background: #0d1117;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 16px 18px;
    overflow-x: auto;
    font-family: var(--mono);
    font-size: 12.5px;
    line-height: 1.7;
    margin: 10px 0;
  }
  .cm  { color: #6e7681; }   /* comment */
  .ky  { color: #ff7b72; }   /* key */
  .st  { color: #a5d6ff; }   /* string */
  .nm  { color: #79c0ff; }   /* number */
  .gn  { color: var(--green); }  /* green */

  /* ── Config table ── */
  .ctable {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    margin: 12px 0;
  }
  .ctable th {
    background: var(--surface2);
    color: var(--accent);
    text-align: left;
    padding: 9px 14px;
    border: 1px solid var(--border);
    font-weight: 600;
    font-size: 12px;
  }
  .ctable td {
    padding: 9px 14px;
    border: 1px solid var(--border);
    vertical-align: top;
  }
  .ctable tr:nth-child(even) td { background: #0d1117; }
  .ctable code { font-size: 12px; }
  .badge {
    display: inline-block;
    padding: 1px 8px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
  }
  .b-red    { background: #2d1a1a; color: var(--red); }
  .b-yellow { background: #2d2206; color: var(--yellow); }
  .b-green  { background: #1a3a2a; color: var(--green); }

  /* ── Note ── */
  .note {
    background: #111d2e;
    border-left: 3px solid var(--accent);
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    font-size: 13.5px;
    color: var(--muted);
    margin: 10px 0;
  }
  .note strong { color: var(--text); }

  .warn {
    background: #2d2206;
    border-left: 3px solid var(--yellow);
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    font-size: 13.5px;
    color: var(--muted);
    margin: 10px 0;
  }

  /* ── Trigger box ── */
  .trigger-row {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
    margin: 14px 0;
  }
  .tbox {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 8px 16px;
    font-size: 13.5px;
    font-weight: 600;
  }
  .tarr { color: var(--accent); font-size: 20px; }

  @media (max-width: 600px) {
    .sheet { grid-template-columns: 1fr; }
    .flow { flex-direction: column; }
    .flow-arrow { transform: rotate(90deg); justify-content: center; }
  }
</style>
</head>
<body>

<!-- ── Hero ── -->
<div class="hero">
  <div class="hero-label">Google Apps Script</div>
  <h1>📊 STOCK_PRICE_DATABASE</h1>
  <p>Setup Guide ฉบับสมบูรณ์ — อัปเดต 2026-06-24</p>
</div>

<!-- ── Flow ── -->
<h2>ภาพรวมระบบ</h2>
<div class="flow">
  <div class="flow-node accent">
    <div class="fn">ไฟล์หลัก</div>
    <div class="label">01_Transaction_Log</div>
    <div class="sub">Backend_ATR · ATR_Portfolio US/TH</div>
  </div>
  <div class="flow-arrow">⇅</div>
  <div class="flow-node green">
    <div class="fn">ไฟล์ Database (ใหม่)</div>
    <div class="label">STOCK_PRICE_DATABASE</div>
    <div class="sub">Daily_Close_Log · Highest_Close_Summary · ATR_History</div>
  </div>
  <div class="flow-arrow">→</div>
  <div class="flow-node yellow">
    <div class="fn">แจ้งเตือน</div>
    <div class="label">Telegram</div>
    <div class="sub">ทุกขั้นตอน</div>
  </div>
</div>

<!-- ── Sheets ── -->
<h2>Sheets ที่สร้างอัตโนมัติ</h2>
<div class="sheet-grid">

  <div class="sheet">
    <div class="sheet-name">Daily_Close_Log</div>
    <div class="sheet-info">
      <div class="sheet-desc">ราคาปิดรายวัน ดึงจาก Yahoo Finance อัตโนมัติ</div>
      <div class="cols">
        <span class="col-tag">A · Date</span>
        <span class="col-tag">B · Symbol</span>
        <span class="col-tag">C · Price</span>
        <span class="col-tag">D · High</span>
        <span class="col-tag">E · Low</span>
        <span class="col-tag">F · Close</span>
      </div>
    </div>
  </div>

  <div class="sheet">
    <div class="sheet-name">Highest_Close_Summary</div>
    <div class="sheet-info">
      <div class="sheet-desc">สถิติราคาปิดสูงสุดตลอดกาล คำนวณอัตโนมัติทุกวัน</div>
      <div class="cols">
        <span class="col-tag">A · Symbol</span>
        <span class="col-tag">B · Highest Close</span>
      </div>
    </div>
  </div>

  <div class="sheet">
    <div class="sheet-name">ATR_History</div>
    <div class="sheet-info">
      <div class="sheet-desc">ATR14 คำนวณด้วย Wilder's Smoothing บันทึกทุกวันอาทิตย์</div>
      <div class="cols">
        <span class="col-tag">A · Date</span>
        <span class="col-tag">B · Symbol</span>
        <span class="col-tag">C · ATR14</span>
      </div>
    </div>
  </div>

</div>

<div class="note" style="margin-top:16px">
  <strong>Ticker ดึงจากไหน?</strong> Script ดึงรายชื่อหุ้นจาก Sheet <code>Backend_ATR</code> ในไฟล์หลัก <code>01_Transaction_Log</code> อัตโนมัติ ไม่ต้องกรอกเองใน Database
</div>

<!-- ── Config ── -->
<h2>ค่าที่ตั้งไว้แล้วในไฟล์</h2>

<table class="ctable">
  <tr><th>ค่า</th><th>ค่าปัจจุบัน</th><th>สถานะ</th></tr>
  <tr>
    <td><code>PORTFOLIO_FILE_ID</code></td>
    <td><code>1nU442ijhr3MBp2dyLJvDIcgqCkfTwlFOiI15SQskdyU</code></td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>BACKEND_ATR_SHEET</code></td>
    <td><code>"Backend_ATR"</code></td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>US_TICKER_COL</code></td>
    <td><code>1</code> (col A)</td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>US_ATR_WRITE_COL</code></td>
    <td><code>2</code> (col B)</td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>TH_TICKER_COL</code></td>
    <td><code>5</code> (col E)</td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>TH_ATR_WRITE_COL</code></td>
    <td><code>6</code> (col F)</td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>ATR_Portfolio US → Highest Close</code></td>
    <td><code>symbolCol:1, writeCol:4</code> (col D)</td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>ATR_Portfolio TH → Highest Close</code></td>
    <td><code>symbolCol:1, writeCol:4</code> (col D)</td>
    <td><span class="badge b-green">✅ ใส่แล้ว</span></td>
  </tr>
  <tr>
    <td><code>TELEGRAM_BOT_TOKEN</code></td>
    <td><code>""</code></td>
    <td><span class="badge b-yellow">ไม่บังคับ</span></td>
  </tr>
</table>

<!-- ── Steps ── -->
<h2>ขั้นตอนติดตั้ง</h2>
<div class="steps">

  <div class="step">
    <div class="step-num">1</div>
    <div class="step-body">
      <div class="step-title">สร้าง Google Sheets ใหม่</div>
      <p>ตั้งชื่อว่า <code>STOCK_PRICE_DATABASE</code> (ชื่ออะไรก็ได้ แต่จำให้ได้)</p>
    </div>
  </div>

  <div class="step">
    <div class="step-num">2</div>
    <div class="step-body">
      <div class="step-title">เปิด Apps Script Editor</div>
      <p>เมนู <code>Extensions</code> → <code>Apps Script</code><br>
      ลบ Code เดิมออกให้หมด แล้ววาง Script จากไฟล์ <code>STOCK_PRICE_DATABASE_Script.js</code></p>
    </div>
  </div>

  <div class="step">
    <div class="step-num">3</div>
    <div class="step-body">
      <div class="step-title">ตรวจสอบ CONFIG ด้านบน (ไม่ต้องแก้ถ้าไม่เปลี่ยน)</div>
      <p>ค่าทุกอย่างตั้งไว้แล้ว ยกเว้น Telegram ถ้าต้องการ:</p>
      <ul>
        <li>แก้ <code>TELEGRAM_BOT_TOKEN</code> และ <code>TELEGRAM_CHAT_ID</code></li>
        <li>ถ้า Highest Close ไม่ได้อยู่ col D ให้แก้ <code>writeCol</code> ใน <code>HIGHEST_CLOSE_TARGETS</code></li>
      </ul>
    </div>
  </div>

  <div class="step">
    <div class="step-num">4</div>
    <div class="step-body">
      <div class="step-title">รัน <code>setupDatabase()</code> ครั้งเดียว</div>
      <p>ใน Apps Script Editor → เลือก function <code>setupDatabase</code> จาก dropdown → กด ▶ Run<br>
      อนุญาต Permission ทั้งหมดที่ Google ขอ → รอ Alert ยืนยัน</p>
      <div class="note" style="margin-top:10px">
        <strong>Alert จะแสดง:</strong> รายชื่อ Ticker ที่ดึงได้จาก Backend_ATR ให้ตรวจสอบว่าครบ
      </div>
    </div>
  </div>

  <div class="step">
    <div class="step-num">5</div>
    <div class="step-body">
      <div class="step-title">ทดสอบรัน <code>runDailyJob()</code> ด้วยมือ</div>
      <p>เปลี่ยน dropdown เป็น <code>runDailyJob</code> → กด ▶ Run<br>
      เปิด <code>View → Logs</code> เพื่อดูผล ควรเห็น:</p>
      <ul>
        <li><code>Tickers US=[VOO, JEPQ, ...] TH=[SCB, KBANK, ...]</code></li>
        <li><code>saveDailyPrice: N รายการ</code></li>
        <li><code>updateHighestClose: N symbols</code></li>
        <li><code>syncPortfolio: เสร็จแล้ว</code></li>
      </ul>
    </div>
  </div>

</div>

<!-- ── Trigger ── -->
<h2>Trigger อัตโนมัติ</h2>
<p style="color:var(--muted);font-size:13.5px;margin-bottom:12px">ตั้งอัตโนมัติตอนรัน setupDatabase() ไม่ต้องทำเอง</p>

<div class="trigger-row">
  <div class="tbox">⏰ ทุกวัน 02:00</div>
  <span class="tarr">→</span>
  <div class="tbox">1 · saveDailyPrice()</div>
  <span class="tarr">→</span>
  <div class="tbox">2 · updateHighestClose()</div>
  <span class="tarr">→</span>
  <div class="tbox">3 · syncPortfolio()</div>
  <span class="tarr">→</span>
  <div class="tbox">4 · calculateATR14() <span style="color:var(--muted);font-weight:400;font-size:12px">(เฉพาะอาทิตย์)</span></div>
</div>

<!-- ── ATR Write Back ── -->
<h2>การเขียนค่ากลับไฟล์หลัก</h2>

<table class="ctable">
  <tr><th>ข้อมูล</th><th>เขียนไปที่</th><th>คอลัมน์</th><th>หมายเหตุ</th></tr>
  <tr>
    <td>ATR14 หุ้นสหรัฐ</td>
    <td><code>Backend_ATR</code></td>
    <td><code>B</code></td>
    <td style="color:var(--muted);font-size:12px">ค่าตัวเลขจริง ทุกวันอาทิตย์</td>
  </tr>
  <tr>
    <td>ATR14 หุ้นไทย</td>
    <td><code>Backend_ATR</code></td>
    <td><code>F</code></td>
    <td style="color:var(--muted);font-size:12px">ค่าตัวเลขจริง ทุกวันอาทิตย์</td>
  </tr>
  <tr>
    <td>Highest Close US</td>
    <td><code>ATR_Portfolio US</code></td>
    <td><code>D</code></td>
    <td style="color:var(--muted);font-size:12px">ค่าตัวเลขจริง ทุกวัน</td>
  </tr>
  <tr>
    <td>Highest Close TH</td>
    <td><code>ATR_Portfolio TH</code></td>
    <td><code>D</code></td>
    <td style="color:var(--muted);font-size:12px">ค่าตัวเลขจริง ทุกวัน</td>
  </tr>
</table>

<div class="warn">
  <strong>ℹ️ ค่าตัวเลขจริง</strong> — Script ใช้ <code>.setValue()</code> ทุกที่ ไม่มีสูตรเขียนลงเซลล์ เซลล์จะไม่เปลี่ยนเมื่อ refresh และไม่หายเมื่อ sheet เปลี่ยนโครงสร้าง
</div>

<!-- ── Highest Close Config ── -->
<h2>เพิ่ม / เปลี่ยน ปลายทาง Highest Close</h2>
<p style="color:var(--muted);font-size:13.5px;margin-bottom:10px">แก้ใน CONFIG → <code>HIGHEST_CLOSE_TARGETS</code> เพิ่ม object ได้ไม่จำกัด</p>

<pre><span class="ky">HIGHEST_CLOSE_TARGETS</span>: [
  <span class="cm">// ปลายทางที่ 1 — ATR_Portfolio US, Highest Close ที่ col D</span>
  { <span class="ky">fileId</span>: <span class="st">""</span>, <span class="ky">sheet</span>: <span class="st">"ATR_Portfolio US"</span>, <span class="ky">symbolCol</span>: <span class="nm">1</span>, <span class="ky">writeCol</span>: <span class="nm">4</span>, <span class="ky">market</span>: <span class="st">"US"</span> },

  <span class="cm">// ปลายทางที่ 2 — ATR_Portfolio TH, Highest Close ที่ col D</span>
  { <span class="ky">fileId</span>: <span class="st">""</span>, <span class="ky">sheet</span>: <span class="st">"ATR_Portfolio TH"</span>, <span class="ky">symbolCol</span>: <span class="nm">1</span>, <span class="ky">writeCol</span>: <span class="nm">4</span>, <span class="ky">market</span>: <span class="st">"TH"</span> },

  <span class="cm">// เพิ่มได้อีกเช่น Backend_ATR col C สำหรับ US</span>
  <span class="cm">// { fileId: "", sheet: "Backend_ATR", symbolCol: 1, writeCol: 3, market: "US" },</span>
],</pre>

<table class="ctable">
  <tr><th>Field</th><th>ความหมาย</th><th>ตัวอย่าง</th></tr>
  <tr><td><code>fileId</code></td><td>ID ไฟล์ปลายทาง — <code>""</code> = ไฟล์หลัก</td><td><code>""</code></td></tr>
  <tr><td><code>sheet</code></td><td>ชื่อ Sheet ปลายทาง</td><td><code>"ATR_Portfolio US"</code></td></tr>
  <tr><td><code>symbolCol</code></td><td>คอลัมน์ที่มี Symbol (A=1, B=2 ...)</td><td><code>1</code></td></tr>
  <tr><td><code>writeCol</code></td><td>คอลัมน์ที่ต้องการเขียน Highest Close</td><td><code>4</code> (col D)</td></tr>
  <tr><td><code>market</code></td><td>กรองหุ้น</td><td><code>"US"</code> / <code>"TH"</code> / <code>"ALL"</code></td></tr>
</table>

<!-- ── Telegram ── -->
<h2>Telegram (ถ้าต้องการ)</h2>
<div class="steps">
  <div class="step">
    <div class="step-num" style="background:var(--yellow)">T</div>
    <div class="step-body">
      <div class="step-title">ขั้นตอนรับ Token และ Chat ID</div>
      <ul>
        <li>เปิด Telegram → ค้นหา <code>@BotFather</code> → พิมพ์ <code>/newbot</code> → คัดลอก Token</li>
        <li>ส่งข้อความหา Bot แล้วเปิด URL นี้:<br>
        <code>https://api.telegram.org/bot&lt;TOKEN&gt;/getUpdates</code></li>
        <li>ดูค่า <code>chat.id</code> ในผลลัพธ์ นำไปใส่ <code>TELEGRAM_CHAT_ID</code></li>
      </ul>
    </div>
  </div>
</div>

<!-- ── Debug ── -->
<h2>Debug / ตรวจสอบ</h2>
<div class="note">
  <strong>ดู Log:</strong> Apps Script Editor → <code>View</code> → <code>Logs</code><br>
  <strong>ดู History:</strong> เมนู <code>Executions</code> — เห็นว่า Trigger รันสำเร็จหรือ Error<br>
  <strong>ทดสอบเอง:</strong> รัน <code>runDailyJob()</code> ด้วยมือได้ตลอดเวลา
</div>

</body>
</html>



## watchlist_backfill.gs
/**
 * ============================================================
 * watchlist_backfill.gs
 * ------------------------------------------------------------
 * 📌 วางไฟล์นี้ใน "โปรเจกต์ STOCK_PRICE_DATABASE" (โปรเจกต์เดียวกับ
 * backfillHistoricalPrices() / saveDailyPrice() / updateHighestClose() /
 * getTickers() / openFile() ใน STOCK_PRICE.gs) — ไม่ใช่โปรเจกต์ webapp_XX.gs
 * เพราะอยู่คนละไฟล์ Apps Script กัน (คนละ global scope) เรียกข้ามกันไม่ได้
 
 * ทำอะไร: ของเดิม backfillHistoricalPrices()/saveDailyPrice() ดึงเฉพาะ Ticker
 * จาก Backend_ATR (หุ้นที่ถืออยู่จริง) ไฟล์นี้เพิ่มการ backfill ย้อนหลัง 6 เดือน
 * ให้ "หุ้นใน Watchlist ที่ยังไม่ได้ถือ" ด้วย แล้วรวมเข้า Daily_Close_Log +
 * Highest_Close_Summary ชุดเดียวกัน (ไฟล์ Database เดียวกับที่
 * webapp_09_external_history.gs เปิดผ่าน EXTERNAL_LOG_SHEET_ID)
 * → ทำให้ getSupportLevels()/getResistanceLevels() หาแนวรับ/แนวต้านให้หุ้นที่
 * ยังไม่ถือได้ด้วย ตามที่ต้องการ
 *
 * ✅ อัปเดตตาม webapp_07_watchlist.gs จริงแล้ว: ชื่อชีต "⭐ Watchlist" (มี emoji)
 * ข้อมูลเริ่มแถว 7 — A=Ticker B=Market(US/TH) ... K=สถานะ(watchlist/cancel)
 *
 * ⚠️ ไฟล์ที่มีชีต ⭐ Watchlist อยู่คนละไฟล์กับ CONFIG.PORTFOLIO_FILE_ID (ไฟล์
 * Backend_ATR/ATR_Portfolio) — ต้องใส่ WATCHLIST_FILE_ID ด้านล่างแยกต่างหาก
 * 🔴 ยังไม่ได้ใส่ ID จริง — ต้องแก้ก่อนใช้งาน ไม่งั้น _wlGetUnheldWatchlistTickers()
 *    จะหาชีตไม่เจอและ log ว่า "ไม่พบชีต" เฉยๆ (ไม่ error แต่ก็ไม่ backfill อะไร)
 * ============================================================
 */

const WATCHLIST_FILE_ID = '1PWw7KfJIKmr1K7f1T24TWaHowHQfGlzzev5wMtS9ffo'; // 🔴 ต้องแก้ก่อนใช้งาน — เอาจาก URL ไฟล์ที่มีชีต ⭐ Watchlist
const WATCHLIST_SHEET_NAME = '⭐ Watchlist'; // ต้องตรงกับชื่อแท็บจริงเป๊ะๆ รวม emoji
const WATCHLIST_START_ROW = 7;
const WATCHLIST_SYMBOL_COL = 1;              // col A
const WATCHLIST_MARKET_COL = 2;              // col B ('TH' | 'US')
const WATCHLIST_STATUS_COL = 11;             // col K ('watchlist' | 'cancel')
const WATCHLIST_BACKFILL_MONTHS = 6;

/**
 * อ่านรายชื่อหุ้นจากชีต Watchlist ในไฟล์ Portfolio หลัก (CONFIG.PORTFOLIO_FILE_ID
 * ผ่าน openFile() เดิม) — คืนเฉพาะ symbol ที่ "ยังไม่ถือ" (ไม่อยู่ใน Backend_ATR
 * US/TH อยู่แล้ว) เพื่อไม่ยิง Yahoo ซ้ำกับที่ runDailyJob() ทำทุกวันอยู่แล้ว
 */
function _wlGetUnheldWatchlistTickers() {
  if (!WATCHLIST_FILE_ID || WATCHLIST_FILE_ID === 'PASTE_WATCHLIST_SPREADSHEET_ID_HERE') {
    Logger.log('_wlGetUnheldWatchlistTickers: ยังไม่ได้ตั้งค่า WATCHLIST_FILE_ID — ใส่ Spreadsheet ID ของไฟล์ที่มีชีต ' + WATCHLIST_SHEET_NAME + ' ก่อน');
    return { us: [], th: [] };
  }
  const watchlistFile = SpreadsheetApp.openById(WATCHLIST_FILE_ID); // คนละไฟล์กับ PORTFOLIO_FILE_ID
  const sheet = watchlistFile.getSheetByName(WATCHLIST_SHEET_NAME);
  if (!sheet) {
    Logger.log('_wlGetUnheldWatchlistTickers: ไม่พบชีต ' + WATCHLIST_SHEET_NAME + ' ในไฟล์ WATCHLIST_FILE_ID');
    return { us: [], th: [] };
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < WATCHLIST_START_ROW) return { us: [], th: [] };

  const numRows = lastRow - WATCHLIST_START_ROW + 1;
  const values = sheet.getRange(WATCHLIST_START_ROW, 1, numRows, WATCHLIST_STATUS_COL).getValues();

  const { us: heldUS, th: heldTH } = getTickers(); // เดิม — จาก Backend_ATR
  const heldUSSet = new Set(heldUS), heldTHSet = new Set(heldTH);

  const us = [], th = [];
  values.forEach(r => {
    const sym = String(r[WATCHLIST_SYMBOL_COL - 1] || '').trim().toUpperCase();
    const market = String(r[WATCHLIST_MARKET_COL - 1] || '').trim().toUpperCase();
    const status = String(r[WATCHLIST_STATUS_COL - 1] || 'watchlist').trim().toLowerCase();
    if (!sym || status === 'cancel') return; // ข้ามแถวว่างและที่ถูกลบ (soft delete)
    if (market === 'TH') {
      if (!heldTHSet.has(sym)) th.push(sym);
    } else {
      if (!heldUSSet.has(sym)) us.push(sym);
    }
  });
  return { us, th };
}

/**
 * ดึง OHLC ย้อนหลังจาก Yahoo Finance — logic เดียวกับ fetchHistory() ที่ซ่อนอยู่
 * ข้างใน backfillHistoricalPrices() เดิม (เรียกจากข้างนอกไม่ได้เพราะเป็น nested
 * function ในฟังก์ชันนั้น) เขียนแยกไว้ที่นี่แทน แต่ใช้ field/สูตรเดียวกันทุกจุด
 * เพื่อให้ผลลัพธ์ตรงกับของเดิม 100%
 */
function _wlFetchYahooOHLC(symbol, period1, period2, existingKeys) {
  try {
    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}`
      + `?interval=1d&period1=${period1}&period2=${period2}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const result = json.chart.result[0];
    const timestamps = result.timestamp;
    const quotes = result.indicators.quote[0];

    const rows = [];
    timestamps.forEach((ts, i) => {
      const close = quotes.close[i];
      const high = quotes.high[i];
      const low = quotes.low[i];
      if (!close) return; // วันหยุดตลาด
      const d = new Date(ts * 1000);
      const cleanSymbol = symbol.replace('.BK', '');
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const key = dateStr + '_' + cleanSymbol;
      if (existingKeys.has(key)) return; // มีแล้ว ข้าม
      rows.push([dateStr, cleanSymbol, close, high, low, close]);
    });
    return rows;
  } catch (e) {
    Logger.log('_wlFetchYahooOHLC error [' + symbol + ']: ' + e);
    return [];
  }
}

/**
 * MAIN — รันด้วยมือครั้งเดียว (ไม่ผูกกับ Trigger รายวันเดิม เพราะถ้า Watchlist
 * มีหุ้นเยอะ การ backfill 6 เดือนทีเดียวจะช้ากว่างานประจำวันปกติมาก)
 * Backfill ราคาย้อนหลัง 6 เดือนให้หุ้น Watchlist ที่ยังไม่ถือ แล้วอัปเดต
 * Highest_Close_Summary ด้วย updateHighestClose() เดิม (คำนวณจาก
 * Daily_Close_Log ทั้งชีต ครอบคลุมของเดิม + ที่เพิ่งเพิ่มอัตโนมัติ)
 */
function backfillWatchlistHistory() {
  const db = getDB();
  const logSheet = getOrCreateSheet(db, 'Daily_Close_Log');
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['Date', 'Symbol', 'Price', 'High', 'Low', 'Close']);
    logSheet.setFrozenRows(1);
  }

  const { us, th } = _wlGetUnheldWatchlistTickers();
  Logger.log(`Watchlist Backfill — US: [${us}] | TH: [${th}]`);
  if (us.length === 0 && th.length === 0) {
    Logger.log('backfillWatchlistHistory: ไม่มีหุ้น Watchlist ที่ยังไม่ถือ (หรือไม่พบชีต Watchlist) — ข้าม');
    writeLog('ข้อมูล', 'backfillWatchlistHistory: ไม่มีหุ้นให้ backfill — ข้าม');
    return;
  }

  const today = new Date();
  const sixAgo = new Date(today);
  sixAgo.setMonth(sixAgo.getMonth() - WATCHLIST_BACKFILL_MONTHS);
  const period1 = Math.floor(sixAgo.getTime() / 1000);
  const period2 = Math.floor(today.getTime() / 1000);

  const existing = new Set();
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues()
      .forEach(r => existing.add(`${r[0]}_${r[1]}`));
  }

  let allRows = [];
  us.forEach(sym => {
    const rows = _wlFetchYahooOHLC(sym, period1, period2, existing);
    Logger.log(`${sym}: ${rows.length} วัน`);
    allRows = allRows.concat(rows);
    Utilities.sleep(500); // กัน rate limit — เหมือนของเดิม
  });
  th.forEach(sym => {
    const rows = _wlFetchYahooOHLC(sym + '.BK', period1, period2, existing);
    Logger.log(`${sym}: ${rows.length} วัน`);
    allRows = allRows.concat(rows);
    Utilities.sleep(500);
  });

  allRows.sort((a, b) => a[0].localeCompare(b[0]));

  if (allRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, allRows.length, 6).setValues(allRows);
    Logger.log(`backfillWatchlistHistory เสร็จ: บันทึก ${allRows.length} แถว`);
  } else {
    Logger.log('backfillWatchlistHistory: ไม่มีข้อมูลใหม่ (มีครบอยู่แล้ว)');
  }

  updateHighestClose(); // เดิม — คำนวณใหม่จาก Daily_Close_Log ทั้งชีต ครอบคลุม watchlist ที่เพิ่งเพิ่ม

  const symbolCount = us.length + th.length;
  writeLog('ข้อมูล', `backfillWatchlistHistory: ${symbolCount} symbols (ยังไม่ถือ) | ${allRows.length} แถวใหม่`);
  sendTelegram(
    `📥 <b>Watchlist Backfill เสร็จแล้ว</b>\n` +
    `📅 ย้อนหลัง ${WATCHLIST_BACKFILL_MONTHS} เดือน\n` +
    `📊 ${symbolCount} Symbols (ที่ยังไม่ถือ) | ${allRows.length} แถวใหม่\n` +
    `✅ Highest Close อัปเดตแล้ว`
  );

  try {
    SpreadsheetApp.getUi().alert(
      `✅ Watchlist Backfill เสร็จแล้ว!\n\n` +
      `📊 Symbols: ${symbolCount} ตัว (เฉพาะที่ยังไม่ถือ)\n` +
      `📋 บันทึกใหม่: ${allRows.length} แถว\n` +
      `💡 Highest_Close_Summary อัปเดตแล้ว`
    );
  } catch (e) { /* รันจาก Trigger — ไม่มี UI */ }
}



## webapp_13_new_ticker_backfill.gs
// ============================================================
// webapp_13_new_ticker_backfill.gs — เฟส 4: หุ้นที่เพิ่งซื้อ (ไฟล์ Database)
// ============================================================
// ไฟล์นี้อยู่ใน STOCK_PRICE_DATABASE (คนละไฟล์กับ Stock_Database/เว็บแอป)
// ใช้ฟังก์ชันเดิมที่มีอยู่แล้วในไฟล์นี้ร่วมกัน: getDB, getOrCreateSheet,
// formatDate, getTickers, openFile, sendTelegram, writeLog — ไม่ต้องประกาศซ้ำ
//
// วิธีทำงาน: เทียบรายชื่อหุ้นใน Backend_ATR (ไฟล์หลัก ผ่าน getTickers() เดิม)
// กับ symbol ที่มีอยู่แล้วใน Daily_Close_Log (ไฟล์นี้) — ตัวไหน "ใหม่"
// (ยังไม่เคยมีข้อมูลราคาเลย) จะดึงย้อนหลัง 6 เดือนเฉพาะตัวนั้น แล้วคำนวณ
// Highest Close / Highest Date / Current Drawdown ใหม่ทั้งชีต
// ============================================================

/**
 * ดึงรายชื่อ symbol ทั้งหมดที่มีอยู่แล้วใน Daily_Close_Log (unique)
 */
function getExistingDailyLogSymbols() {
  const db = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  const existing = new Set();
  if (!logSheet || logSheet.getLastRow() < 2) return existing;

  logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues()
    .forEach(r => {
      const sym = String(r[0]).trim();
      if (sym) existing.add(sym);
    });
  return existing;
}

/**
 * ดึงราคาย้อนหลัง (period1→period2) ของ 1 symbol จาก Yahoo Finance
 * แยกออกมาจาก backfillHistoricalPrices() เดิม เพื่อใช้ซ้ำได้ทั้งแบบ
 * backfill ทั้งพอร์ต (ของเดิม) และ backfill เฉพาะหุ้นใหม่ (เฟส 4)
 */
function fetchSymbolHistory(symbol, period1, period2) {
  try {
    const url  = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}`
               + `?interval=1d&period1=${period1}&period2=${period2}`;
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const result = json.chart.result[0];

    const timestamps = result.timestamp;
    const quotes = result.indicators.quote[0];
    const cleanSymbol = symbol.replace(".BK", "");

    const rows = [];
    timestamps.forEach((ts, i) => {
      const close = quotes.close[i];
      const high  = quotes.high[i];
      const low   = quotes.low[i];
      if (!close) return; // วันหยุดตลาด

      const d = new Date(ts * 1000);
      rows.push([
        Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        cleanSymbol,
        close, high, low, close,
      ]);
    });
    return rows;
  } catch (e) {
    Logger.log(`fetchSymbolHistory error [${symbol}]: ${e}`);
    return [];
  }
}

/**
 * เฟส 4 — ตรวจหุ้นที่เพิ่งซื้อ (มีใน Backend_ATR แต่ยังไม่มีข้อมูลราคาใน
 * Daily_Close_Log เลย) แล้ว backfill ย้อนหลัง 6 เดือนเฉพาะตัวนั้น
 * รันเองด้วยมือได้ หรือใส่เป็น Step 0 ใน runDailyJob() (ดูท้ายไฟล์)
 */
function backfillNewTickers() {
  const db = getDB();
  const logSheet = getOrCreateSheet(db, "Daily_Close_Log");
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Date", "Symbol", "Price", "High", "Low", "Close"]);
    logSheet.setFrozenRows(1);
  }

  const { us, th } = getTickers();
  const existing = getExistingDailyLogSymbols();

  const newUS = us.filter(sym => !existing.has(sym));
  const newTH = th.filter(sym => !existing.has(sym));

  if (newUS.length === 0 && newTH.length === 0) {
    writeLog("ข้อมูล", "backfillNewTickers: ไม่มีหุ้นใหม่ที่ต้อง backfill");
    return { newCount: 0, symbols: [] };
  }

  const today  = new Date();
  const sixAgo = new Date(today);
  sixAgo.setMonth(sixAgo.getMonth() - 6);
  const period1 = Math.floor(sixAgo.getTime() / 1000);
  const period2 = Math.floor(today.getTime() / 1000);

  let allRows = [];

  newUS.forEach(sym => {
    const rows = fetchSymbolHistory(sym, period1, period2);
    Logger.log(`[ใหม่] ${sym}: ${rows.length} วัน`);
    allRows = allRows.concat(rows);
    Utilities.sleep(500);
  });

  newTH.forEach(sym => {
    const rows = fetchSymbolHistory(sym + ".BK", period1, period2);
    Logger.log(`[ใหม่] ${sym}: ${rows.length} วัน`);
    allRows = allRows.concat(rows);
    Utilities.sleep(500);
  });

  allRows.sort((a, b) => a[0].localeCompare(b[0]));

  if (allRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, allRows.length, 6).setValues(allRows);
  }

  // คำนวณ Highest Close / Date / Drawdown ใหม่ทั้งชีต (หุ้นเก่าก็ได้ประโยชน์ด้วย
  // เพราะ Current Price อาจเปลี่ยนจาก saveDailyPrice() รอบล่าสุดเช่นกัน)
  updateHighestClose();
  syncPortfolio();

  const allNew = [...newUS, ...newTH];
  const msg = `backfillNewTickers: พบหุ้นใหม่ ${allNew.length} ตัว [${allNew.join(", ")}] · บันทึก ${allRows.length} แถว`;
  Logger.log(msg);
  writeLog("ข้อมูล", msg);
  sendTelegram(
    `🆕 <b>หุ้นใหม่ Backfill เสร็จแล้ว</b>\n` +
    `📊 ${allNew.length} ตัว: ${allNew.join(", ")}\n` +
    `📋 บันทึก ${allRows.length} แถว (ย้อนหลัง 6 เดือน)\n` +
    `✅ Highest Close / Drawdown อัปเดตแล้ว`
  );

  return { newCount: allNew.length, symbols: allNew };
}

// ============================================================
// updateHighestClose() เวอร์ชันใหม่ — แทนที่ตัวเดิมทั้งฟังก์ชัน
// เดิมมี Symbol / Highest Close / Date เฉยๆ — เพิ่ม Current Price +
// Current Drawdown % ตามสเปคเฟส 4
// ⚠️ ลบฟังก์ชัน updateHighestClose() เดิมออกก่อน แล้ววางตัวนี้แทนที่
// (syncPortfolio() เดิมไม่ต้องแก้ เพราะยังอ่านแค่ col A/B เท่าเดิม)
// ============================================================
function updateHighestClose() {
  const db = getDB();
  const logSheet = db.getSheetByName("Daily_Close_Log");
  if (!logSheet || logSheet.getLastRow() < 2) return;

  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();

  const highest = {}; // { sym: { value, date } }
  const latest  = {}; // { sym: { close, date } } — แถวล่าสุดของแต่ละ symbol

  data.forEach(row => {
    const dateVal = row[0];
    const dateStr = dateVal instanceof Date ? formatDate(dateVal) : String(dateVal).substring(0, 10);
    const sym   = String(row[1]).trim();
    const close = parseFloat(row[5]);
    if (!sym || isNaN(close)) return;

    if (!highest[sym] || close > highest[sym].value) {
      highest[sym] = { value: close, date: dateStr };
    }
    if (!latest[sym] || dateStr > latest[sym].date) {
      latest[sym] = { close: close, date: dateStr };
    }
  });

  const sumSheet = getOrCreateSheet(db, "Highest_Close_Summary");
  sumSheet.clearContents();
  sumSheet.appendRow(["Symbol", "Highest Close", "Highest Date", "Current Price", "Current Drawdown %"]);
  sumSheet.setFrozenRows(1);
  sumSheet.getRange(1, 1, 1, 5).setFontWeight("bold");

  Object.keys(highest).sort().forEach(sym => {
    const h = highest[sym];
    const l = latest[sym];
    const drawdownPct = l ? Math.round(((l.close - h.value) / h.value) * 10000) / 100 : "";
    sumSheet.appendRow([sym, h.value, h.date, l ? l.close : "", drawdownPct]);
  });

  if (sumSheet.getLastRow() > 1) {
    sumSheet.getRange(2, 3, sumSheet.getLastRow() - 1, 1).setNumberFormat("yyyy-mm-dd");
    sumSheet.getRange(2, 5, sumSheet.getLastRow() - 1, 1).setNumberFormat('0.00"%"');
  }

  Logger.log(`updateHighestClose: ${Object.keys(highest).length} symbols (Date/Price/Drawdown)`);
}

// ============================================================
// วิธี HOOK เข้ากับ runDailyJob() เดิม — เพิ่ม Step 0 ก่อน saveDailyPrice()
// ============================================================
//
// function runDailyJob() {
//   writeLog("ประกาศ", "เริ่มการดำเนินการแล้ว");
//   try {
//     writeLog("ข้อมูล", "=== START runDailyJob ===");
//
//     // ── Step 0 : backfillNewTickers (ใหม่ — เช็คหุ้นที่เพิ่งซื้อ) ──
//     writeLog("ข้อมูล", "เริ่ม backfillNewTickers()...");
//     backfillNewTickers();
//
//     // ── Step 1 : saveDailyPrice ── (โค้ดเดิมทั้งหมด ไม่ต้องแก้อะไรต่อ)
//     ...
//
// ทางเลือก: ไม่อยากรอรอบ 02:00 ของวันถัดไป ก็รัน backfillNewTickers() ด้วยมือ
// ได้ทันทีจาก Apps Script Editor (เลือก function จาก dropdown → ▶ Run)
// ============================================================

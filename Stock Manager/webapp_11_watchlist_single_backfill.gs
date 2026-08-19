/**
 * ============================================================
 * webapp_11_watchlist_single_backfill.gs
 * ------------------------------------------------------------
 * Backfill ราคาย้อนหลัง 6 เดือนให้หุ้น Watchlist ทีละตัว — เรียกจากปุ่มบน
 * การ์ดหุ้นแต่ละตัวในหน้า Watchlist โดยตรง:
 *   google.script.run.backfillSingleWatchlistTicker(ticker, market)
 *
 * ⚠️ ทำไมแยกจาก backfillWatchlistHistory() ที่ส่งไปก่อนหน้า:
 * ฟังก์ชันนั้นอยู่ในโปรเจกต์ STOCK_PRICE_DATABASE (คนละไฟล์ Apps Script)
 * ส่วนนี้คือ webapp — คนละ global scope กัน เรียกข้ามโปรเจกต์ผ่าน
 * google.script.run ไม่ได้ ไฟล์นี้จึงเขียนเข้า Daily_Close_Log /
 * Highest_Close_Summary "ตรงๆ" ผ่าน SpreadsheetApp.openById() แทน — ใช้
 * EXTERNAL_LOG_SHEET_ID / EXTERNAL_LOG_SHEET_NAME / EXTERNAL_LOG_COL /
 * HIGHEST_CLOSE_SHEET_NAME / HIGHEST_CLOSE_COL ตัวเดียวกับที่ประกาศไว้แล้ว
 * ใน webapp_09_external_history.gs (โปรเจกต์เดียวกัน ไม่ต้องประกาศซ้ำ)
 *
 * ต่างจากตัว batch ตรงที่ทำทีละ ticker เดียว — ไม่ recompute
 * Highest_Close_Summary ทั้งชีต แค่แถวของ ticker นี้เท่านั้น (เร็วกว่า
 * ปลอดภัยกว่า ไม่แตะข้อมูลหุ้นตัวอื่น)
 * ============================================================
 */

const WL_SINGLE_BACKFILL_MONTHS = 6;

function backfillSingleWatchlistTicker(ticker, market) {
  try {
    if (!EXTERNAL_LOG_SHEET_ID || EXTERNAL_LOG_SHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
      return { success: false, error: 'ยังไม่ได้ตั้งค่า EXTERNAL_LOG_SHEET_ID ใน webapp_09_external_history.gs' };
    }
    ticker = String(ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'กรุณาระบุ ticker' };

    const isTH = String(market || '').trim().toUpperCase() === 'TH';
    const yahooSymbol = isTH ? ticker + '.BK' : ticker;

    const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);
    const logSheet = extSS.getSheetByName(EXTERNAL_LOG_SHEET_NAME) || extSS.insertSheet(EXTERNAL_LOG_SHEET_NAME);
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Date', 'Symbol', 'Price', 'High', 'Low', 'Close']);
      logSheet.setFrozenRows(1);
    }

    // ── กันซ้ำ: เอาเฉพาะวันที่ที่มีอยู่แล้วของ ticker นี้ ──
    const existing = new Set();
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.getRange(2, 1, lastRow - 1, 6).getValues().forEach(r => {
        const sym = String(r[EXTERNAL_LOG_COL.SYMBOL - 1] || '').trim().toUpperCase();
        if (sym !== ticker) return;
        const raw = r[EXTERNAL_LOG_COL.DATE - 1];
        const dateStr = raw instanceof Date
          ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(raw).substring(0, 10);
        existing.add(dateStr);
      });
    }

    const today = new Date();
    const monthsAgo = new Date(today);
    monthsAgo.setMonth(monthsAgo.getMonth() - WL_SINGLE_BACKFILL_MONTHS);
    const period1 = Math.floor(monthsAgo.getTime() / 1000);
    const period2 = Math.floor(today.getTime() / 1000);

    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooSymbol)}`
      + `?interval=1d&period1=${period1}&period2=${period2}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result) {
      return { success: false, error: 'ไม่พบข้อมูลราคาจาก Yahoo Finance สำหรับ ' + yahooSymbol + ' (เช็คว่าพิมพ์ ticker/ตลาดถูกไหม)' };
    }

    const timestamps = result.timestamp || [];
    const quotes = result.indicators.quote[0];

    const rows = [];
    let highestClose = null, highestDate = null;
    timestamps.forEach((ts, i) => {
      const close = quotes.close[i];
      const high = quotes.high[i];
      const low = quotes.low[i];
      if (!close) return; // วันหยุดตลาด
      const d = new Date(ts * 1000);
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (!existing.has(dateStr)) {
        rows.push([dateStr, ticker, close, high, low, close]);
      }
      if (highestClose === null || close > highestClose) { highestClose = close; highestDate = dateStr; }
    });

    if (rows.length > 0) {
      rows.sort((a, b) => a[0].localeCompare(b[0]));
      logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    }

    if (highestClose !== null) {
      _wlUpdateHighestCloseSingle(extSS, ticker, highestClose, highestDate);
    }

    return {
      success: true, ticker, addedRows: rows.length,
      totalDaysAvailable: existing.size + rows.length,
      highestClose, highestDate
    };
  } catch (e) {
    logError('backfillSingleWatchlistTicker', e);
    return { success: false, error: e.message };
  }
}

/**
 * อัปเดต Highest_Close_Summary เฉพาะแถวของ ticker นี้ — ไม่แตะแถวหุ้นตัวอื่น
 * ใช้ HIGHEST_CLOSE_SHEET_NAME/HIGHEST_CLOSE_COL เดิมจาก webapp_09
 */
function _wlUpdateHighestCloseSingle(extSS, ticker, highestClose, highestDate) {
  const sheet = extSS.getSheetByName(HIGHEST_CLOSE_SHEET_NAME) || extSS.insertSheet(HIGHEST_CLOSE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Symbol', 'Highest Close', 'Date']);
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  let rowIdx = -1;
  let currentHighest = -Infinity;
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][HIGHEST_CLOSE_COL.SYMBOL - 1] || '').trim().toUpperCase() === ticker) {
        rowIdx = 2 + i;
        currentHighest = Number(values[i][HIGHEST_CLOSE_COL.HIGHEST_CLOSE - 1]) || -Infinity;
        break;
      }
    }
  }

  if (rowIdx === -1) {
    sheet.appendRow([ticker, highestClose, highestDate]);
    sheet.getRange(sheet.getLastRow(), HIGHEST_CLOSE_COL.DATE).setNumberFormat('yyyy-mm-dd');
  } else if (highestClose > currentHighest) {
    sheet.getRange(rowIdx, HIGHEST_CLOSE_COL.HIGHEST_CLOSE).setValue(highestClose);
    sheet.getRange(rowIdx, HIGHEST_CLOSE_COL.DATE).setValue(highestDate).setNumberFormat('yyyy-mm-dd');
  }
}


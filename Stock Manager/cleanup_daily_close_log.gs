// ============================================================
// cleanup_daily_close_log.gs — ลบราคาหุ้นเก่าใน Daily_Close_Log (ไฟล์ภายนอก)
// เงื่อนไข: ticker ที่ไม่มีทั้งใน Holdings (US/TH) และ Watchlist (active)
// จะถูกลบทิ้ง กันไฟล์บวมจากหุ้นที่ขายหมด/เอาออกจาก Watchlist ไปนานแล้ว
// ⚠️ ใช้ EXTERNAL_LOG_SHEET_ID / EXTERNAL_LOG_SHEET_NAME / EXTERNAL_LOG_COL
//    ตัวเดียวกับ webapp_09_external_history.gs
// ============================================================

// ── รวบรวม ticker ที่ "ต้องเก็บไว้" — ถือครองอยู่ ( เพิ่มเงื่อนไข `sharesRemain > 0) + Watchlist ที่ยัง active ──
function _getTickersToKeep() {
  const keep = new Set();

  try {
    const holdings = getHoldingsData();
    (holdings.us || []).forEach(r => {
      const qty = parseFloat(r.sharesRemain);
      if (r.ticker && qty > 0) keep.add(String(r.ticker).trim().toUpperCase()); // ← เพิ่มเงื่อนไข qty > 0
    });
    (holdings.th || []).forEach(r => {
      const qty = parseFloat(r.sharesRemain);
      if (r.ticker && qty > 0) keep.add(String(r.ticker).trim().toUpperCase()); // ← เพิ่มเงื่อนไข qty > 0
    });
  } catch (e) { logError('_getTickersToKeep:holdings', e); }

  try {
    const watchlist = getWatchlistData();
    (watchlist.items || []).forEach(r => { if (r.ticker) keep.add(String(r.ticker).trim().toUpperCase()); });
  } catch (e) { logError('_getTickersToKeep:watchlist', e); }

  return keep;
}

// ══════════════════════════════════════════════════════════
// ลบแถวใน Daily_Close_Log ของ ticker ที่ไม่อยู่ใน keep-list
// คืนสรุปผล { success, deletedRows, keptTickers, removedTickers }
// ══════════════════════════════════════════════════════════
function cleanupDailyCloseHighLog() {
  try {
    if (!EXTERNAL_LOG_SHEET_ID || EXTERNAL_LOG_SHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
      throw new Error('ยังไม่ได้ตั้งค่า EXTERNAL_LOG_SHEET_ID');
    }

    const keepSet = _getTickersToKeep();
    if (keepSet.size === 0) {
      logError('cleanupDailyCloseHighLog', new Error('keep-list ว่างเปล่า — ยกเลิกการลบเพื่อความปลอดภัย (กันลบทั้งชีตโดยไม่ตั้งใจถ้า Holdings/Watchlist โหลดพลาด)'));
      return { success: false, error: 'ตรวจไม่พบ ticker ใน Holdings/Watchlist เลย — ยกเลิกการลบเพื่อความปลอดภัย' };
    }

    const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);
    const sheet = extSS.getSheetByName(EXTERNAL_LOG_SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีต "' + EXTERNAL_LOG_SHEET_NAME + '"');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, deletedRows: 0, message: 'ไม่มีข้อมูลให้ลบ' };

    const numRows = lastRow - 1;
    const data = sheet.getRange(2, 1, numRows, 6).getValues();

    const removedTickers = new Set();
    const keptRows = [];

    data.forEach(row => {
      const symbol = String(row[EXTERNAL_LOG_COL.SYMBOL - 1] || '').trim().toUpperCase();
      if (!symbol) return; // แถวว่าง ข้ามไป (ไม่เก็บ ไม่นับลบ)
      if (keepSet.has(symbol)) {
        keptRows.push(row);
      } else {
        removedTickers.add(symbol);
      }
    });

    const deletedRows = numRows - keptRows.length;

    if (deletedRows > 0) {
      // ล้างพื้นที่ข้อมูลเดิมทั้งหมดแล้วเขียนแถวที่เหลือกลับเข้าไปใหม่ (เร็วกว่าลบทีละแถว)
      sheet.getRange(2, 1, numRows, 6).clearContent();
      if (keptRows.length > 0) {
        sheet.getRange(2, 1, keptRows.length, 6).setValues(keptRows);
      }
    }

      // ── ลบใน Highest_Close_Summary ด้วย keep-list ชุดเดียวกัน ──
    const highestResult = _cleanupHighestCloseSummary(keepSet);

    return {
      success: true,
      dailyCloseLog: {
        deletedRows,
        keptRows: keptRows.length,
        removedTickers: Array.from(removedTickers)
      },
      highestCloseSummary: highestResult,
      keptTickerCount: keepSet.size
    };
  } catch (e) {
    logError('cleanupDailyCloseHighLog', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// ลบแถวใน Highest_Close_Summary ของ ticker ที่ไม่อยู่ใน keep-list
// ใช้ HIGHEST_CLOSE_SHEET_NAME / HIGHEST_CLOSE_COL ตัวเดียวกับ
// webapp_09_external_history.gs (คอลัมน์ A=Symbol, B=HighestClose, C=Date)
// ══════════════════════════════════════════════════════════
function _cleanupHighestCloseSummary(keepSet) {
  try {
    const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);
    const sheet = extSS.getSheetByName(HIGHEST_CLOSE_SHEET_NAME);
    if (!sheet) return { deletedRows: 0, removedTickers: [], skipped: true, reason: 'ไม่พบชีต ' + HIGHEST_CLOSE_SHEET_NAME };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { deletedRows: 0, removedTickers: [] };

    const numRows = lastRow - 1;
    const data = sheet.getRange(2, 1, numRows, 5).getValues();

    const removedTickers = new Set();
    const keptRows = [];

    data.forEach(row => {
      const symbol = String(row[HIGHEST_CLOSE_COL.SYMBOL - 1] || '').trim().toUpperCase();
      if (!symbol) return;
      if (keepSet.has(symbol)) {
        keptRows.push(row);
      } else {
        removedTickers.add(symbol);
      }
    });

    const deletedRows = numRows - keptRows.length;

    if (deletedRows > 0) {
      sheet.getRange(2, 1, numRows, 5).clearContent();
      if (keptRows.length > 0) {
        sheet.getRange(2, 1, keptRows.length, 5).setValues(keptRows);
      }
    }

    return { deletedRows, removedTickers: Array.from(removedTickers) };
  } catch (e) {
    logError('_cleanupHighestCloseSummary', e);
    return { deletedRows: 0, removedTickers: [], error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// ตั้ง trigger รันอัตโนมัติทุกวัน ตอนตี 3 (หลังราคาปิดตลาด/หลัง backfill รายวัน)
// รันฟังก์ชันนี้ "ครั้งเดียว" จาก Apps Script Editor (เลือกฟังก์ชันนี้ → กด Run)
// เพื่อลงทะเบียน trigger — วางโค้ดอย่างเดียวไม่พอ ต้องกด Run เอง 1 ครั้ง
// ══════════════════════════════════════════════════════════
function createcleanupDailyCloseHighLogTrigger() {
  // ลบ trigger เก่าของฟังก์ชันนี้ก่อน กันสร้างซ้ำถ้าเคยรันมาแล้ว
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'cleanupDailyCloseHighLog') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('cleanupDailyCloseHighLog')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  Logger.log('✅ ตั้ง trigger เรียบร้อย — จะรัน cleanupDailyCloseHighLog() ทุกวันตอนตี 3');
}


// ── ทดสอบดูผลลัพธ์ก่อนตั้ง trigger จริง (ไม่ลบจริง แค่รายงานว่าจะลบอะไรบ้าง) ──
function testcleanupDailyCloseHighLogDryRun() {
  const keepSet = _getTickersToKeep();
  Logger.log('Ticker ที่จะเก็บไว้ (' + keepSet.size + ' ตัว): ' + Array.from(keepSet).join(', '));

  const extSS = SpreadsheetApp.openById(EXTERNAL_LOG_SHEET_ID);

  // ── ตรวจ Daily_Close_Log ──
  const logSheet = extSS.getSheetByName(EXTERNAL_LOG_SHEET_NAME);
  const logLastRow = logSheet.getLastRow();
  if (logLastRow < 2) {
    Logger.log('Daily_Close_Log: ไม่มีข้อมูล');
  } else {
    const logData = logSheet.getRange(2, 1, logLastRow - 1, 6).getValues();
    const willRemoveLog = new Set();
    logData.forEach(row => {
      const symbol = String(row[EXTERNAL_LOG_COL.SYMBOL - 1] || '').trim().toUpperCase();
      if (symbol && !keepSet.has(symbol)) willRemoveLog.add(symbol);
    });
    Logger.log('Daily_Close_Log — Ticker ที่จะถูกลบ (' + willRemoveLog.size + ' ตัว): ' + Array.from(willRemoveLog).join(', '));
  }

  // ── ตรวจ Highest_Close_Summary ──
  const summarySheet = extSS.getSheetByName(HIGHEST_CLOSE_SHEET_NAME);
  if (!summarySheet) {
    Logger.log('Highest_Close_Summary: ไม่พบชีตนี้');
    return;
  }
  const summaryLastRow = summarySheet.getLastRow();
  if (summaryLastRow < 2) {
    Logger.log('Highest_Close_Summary: ไม่มีข้อมูล');
    return;
  }
  const summaryData = summarySheet.getRange(2, 1, summaryLastRow - 1, 3).getValues();
  const willRemoveSummary = new Set();
  summaryData.forEach(row => {
    const symbol = String(row[HIGHEST_CLOSE_COL.SYMBOL - 1] || '').trim().toUpperCase();
    if (symbol && !keepSet.has(symbol)) willRemoveSummary.add(symbol);
  });
  Logger.log('Highest_Close_Summary — Ticker ที่จะถูกลบ (' + willRemoveSummary.size + ' ตัว): ' + Array.from(willRemoveSummary).join(', '));
}

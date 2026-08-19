// ============================================================
// watchlist_auto_update.gs
// อัปเดตราคาปัจจุบันในชีต "⭐ Watchlist" อัตโนมัติผ่าน Trigger
// ------------------------------------------------------------
// เดิม updateWatchlistPricesWeb() (webapp_07_watchlist.gs) ทำงานถูกต้อง
// แต่ถูกเรียกจากปุ่ม "รีเฟรช" ในหน้าเว็บแอปเท่านั้น ไม่มี Trigger อัตโนมัติ
// ทำให้ราคาในคอลัมน์ I ค้างถ้าไม่มีใครเปิดแอป — ไฟล์นี้เพิ่มชั้นห่อ
// (wrapper) ให้รันเป็นเวลาอัตโนมัติ โดย "ไม่แก้" updateWatchlistPricesWeb()
// เดิมเลย (reuse ตรงๆ กันโค้ดซ้ำซ้อน)
//
// วิธีติดตั้ง:
//  1. เพิ่มไฟล์นี้เข้าไปในโปรเจกต์ (ไม่ต้องลบ/แก้ไฟล์เดิม)
//  2. รัน createWatchlistPriceUpdateTrigger() ครั้งเดียว (ตั้ง Trigger)
//  3. (แนะนำ) รัน testUpdateWatchlistPricesAuto() ทดสอบ 1 ครั้งก่อน
// ============================================================

// ── รันอัตโนมัติทุก 30 นาทีผ่าน Trigger ──
// (ตั้งห่างจาก updateThaiPrice/updateUSPriceTrigger ที่รันทุก 15 นาที
//  เพื่อลดโอกาสยิง Yahoo Finance พร้อมกันถี่เกินไปจนโดนบล็อกชั่วคราว)
//
// อัปเดตทั้งคอลัมน์ I (ราคา ณ ปัจจุบัน) และคอลัมน์ J (% เปลี่ยนแปลง —
// เทียบกับราคาตอนเพิ่มเข้า Watchlist ที่คอลัมน์ H) เหมือนกับที่
// updateWatchlistPricesWeb() ทำตอนกดปุ่ม "รีเฟรช" ในหน้าเว็บแอป
function updateWatchlistPricesAuto() {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < WATCHLIST_SHEET.START_ROW) {
      logInfo('updateWatchlistPricesAuto', 'ไม่มีรายการใน Watchlist');
      return;
    }

    const numRows = lastRow - WATCHLIST_SHEET.START_ROW + 1;
    const rows = sheet.getRange(WATCHLIST_SHEET.START_ROW, 1, numRows, 11).getValues();

    let updated = 0;
    const failed = [];

    rows.forEach((row, i) => {
      const ticker = String(row[0] || '').trim();
      const status = String(row[10] || 'watchlist').trim().toLowerCase();
      if (!ticker || status === 'cancel') return;

      const market = String(row[1] || 'US').trim();
      const priceAdded = parseFloat(row[7]) || 0; // col H — ราคาตอนเพิ่มเข้า Watchlist

      const quote = _wlFetchYahooQuote(ticker, market);
      if (!quote) { failed.push(ticker); return; }

      const r = WATCHLIST_SHEET.START_ROW + i;
      sheet.getRange(r, 9).setValue(quote.price); // col I — ราคาปัจจุบัน
      if (priceAdded > 0) {
        sheet.getRange(r, 10).setValue(((quote.price - priceAdded) / priceAdded) * 100); // col J — % เปลี่ยนแปลง
      }
      updated++;
      Utilities.sleep(200);
    });

    logInfo('updateWatchlistPricesAuto',
      'อัปเดตราคา (คอลัมน์ I/J) Watchlist แล้ว ' + updated + ' ตัว' +
      (failed.length ? ' / ล้มเหลว ' + failed.length + ' ตัว' : '')
    );

    // แจ้งเตือนเฉพาะตอนมีตัวที่ดึงราคาไม่ได้ — กันสแปม Telegram ตอนอัปเดตปกติทุก 30 นาที
    if (failed.length) {
      sendTelegramSafe(
        '⚠️ อัปเดตราคา Watchlist บางตัวไม่สำเร็จ\n' +
        '✅ สำเร็จ ' + updated + ' ตัว\n' +
        '❌ ล้มเหลว ' + failed.length + ' ตัว: ' + failed.join(', ') + '\n' +
        '🕐 ' + getNow()
      );
    }
  } catch (e) {
    sendTelegramError('updateWatchlistPricesAuto', e);
  }
}

// ── ตั้ง Trigger รันอัตโนมัติ — รันฟังก์ชันนี้ครั้งเดียวตอนติดตั้ง ──
function createWatchlistPriceUpdateTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateWatchlistPricesAuto')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('updateWatchlistPricesAuto')
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log('createWatchlistPriceUpdateTrigger: ตั้ง Trigger อัปเดตราคา Watchlist ทุก 30 นาทีแล้ว');
  try {
    SpreadsheetApp.getUi().alert('✅ ตั้ง Trigger อัปเดตราคา Watchlist อัตโนมัติแล้ว (ทุก 30 นาที)');
  } catch (e) { /* รันจาก Editor ไม่มี UI ก็ไม่เป็นไร */ }
}

// ── ทดสอบรันด้วยมือ 1 ครั้ง — เช็คคอลัมน์ I/J ในชีต ⭐ Watchlist ว่าอัปเดตจริง ──
function testUpdateWatchlistPricesAuto() {
  updateWatchlistPricesAuto();
}



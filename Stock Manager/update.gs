// ========================================
// update.gs
// ========================================

// ----------------------------------------
// /Updatepcs — Update Price ทั้ง TH และ US
// ----------------------------------------
function updatePrices() {
  try {
    sendTelegramSafe("⏳ กำลัง Update ราคาหุ้น...");

    const thResult = updateThaiPrice();
    Utilities.sleep(1000);
    const usResult = updateUSPrice();

    sendTelegramSafe(
      "✅ Update Price สำเร็จ\n\n" +
      "🇹🇭 หุ้นไทย   : อัปเดต " + thResult + " ตัว\n" +
      "🇺🇸 หุ้นสหรัฐ : อัปเดต " + usResult + " ตัว\n\n" +
      "🕐 " + getNow()
    );
  } catch (e) {
    sendTelegramError("updatePrices", e);
  }
}

// ----------------------------------------
// อัปเดตราคาหุ้นไทย — col G
// ----------------------------------------
function updateThaiPrice() {
  const sheet   = getSheet(SHEETS.TH_HOLD);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.HOLD) return 0;

  const rows = sheet.getRange(
    START_ROW.HOLD, 1,
    lastRow - START_ROW.HOLD + 1, 13
  ).getValues();

  let count = 0;
  rows.forEach((row, i) => {
    const ticker       = row[HOLD_COL.TICKER - 1];
    const sharesRemain = Number(row[HOLD_COL.SHARES_REMAIN - 1]) || 0;
    const symbolFull   = String(row[11] || "").trim();
    if (!ticker || sharesRemain <= 0 || !symbolFull) return;

    const symbol = symbolFull.includes(":")
      ? symbolFull.split(":")[1] + ".BK"
      : symbolFull + ".BK";

    // ✅ ใช้ fetchYahooPrice พร้อม retry
    const data = fetchYahooPrice(symbol);
    if (!data || !data.price) {
      logError("updateThaiPrice", "ไม่พบราคา: " + symbol);
      return;
    }

    sheet.getRange(START_ROW.HOLD + i, HOLD_COL.PRICE_NOW).setValue(data.price);
    count++;
    Utilities.sleep(400);
  });

   logInfo("updateThaiPrice", "อัปเดต " + count + " ตัว");
  _clearHoldingsDataMemo();   // ✅ เขียน PRICE_NOW ลง TH_HOLD แล้ว เคลียร์ memo กัน getHoldingsData() คืนค่าเก่า
  return count;
}


// ----------------------------------------
// อัปเดตราคาหุ้นสหรัฐ — col G
// ----------------------------------------
function updateUSPrice() {
  const sheet   = getSheet(SHEETS.US_HOLD);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.HOLD) return 0;

  const rows = sheet.getRange(
    START_ROW.HOLD, 1,
    lastRow - START_ROW.HOLD + 1, 15
  ).getValues();

  let count = 0;
  rows.forEach((row, i) => {
    const ticker       = row[HOLD_COL.TICKER - 1];
    const sharesRemain = Number(row[HOLD_COL.SHARES_REMAIN - 1]) || 0;
    if (!ticker || sharesRemain <= 0) return;

    // ✅ ใช้ fetchYahooPrice พร้อม retry
    const data = fetchYahooPrice(String(ticker).trim().toUpperCase());
    if (!data || !data.price) {
      logError("updateUSPrice", "ไม่พบราคา: " + ticker);
      return;
    }

    sheet.getRange(START_ROW.HOLD + i, HOLD_COL.PRICE_NOW).setValue(data.price);
    count++;
    Utilities.sleep(400);
  });

   logInfo("updateUSPrice", "อัปเดต " + count + " ตัว");
  _clearHoldingsDataMemo();   // ✅ เขียน PRICE_NOW ลง US_HOLD แล้ว เคลียร์ memo กัน getHoldingsData() คืนค่าเก่า
  return count;
}


// ----------------------------------------
// อัปเดตราคาหุ้นตัวเดียว (เฉพาะ ticker ที่ระบุ)
// ----------------------------------------
function updateSingleThaiPrice(ticker) {
  ticker = String(ticker || '').trim().toUpperCase();
  if (!ticker) return false;

  const sheet   = getSheet(SHEETS.TH_HOLD);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.HOLD) return false;

  const numRows = lastRow - START_ROW.HOLD + 1;
  const rows = sheet.getRange(START_ROW.HOLD, 1, numRows, 13).getValues();

  for (let i = 0; i < rows.length; i++) {
    const rowTicker = String(rows[i][HOLD_COL.TICKER - 1] || '').trim().toUpperCase();
    if (rowTicker !== ticker) continue;

    const sharesRemain = Number(rows[i][HOLD_COL.SHARES_REMAIN - 1]) || 0;
    const symbolFull   = String(rows[i][11] || '').trim();
    if (sharesRemain <= 0) return false;

    const symbol = symbolFull
      ? (symbolFull.includes(':') ? symbolFull.split(':')[1] + '.BK' : symbolFull + '.BK')
      : ticker + '.BK';

    const data = fetchYahooPrice(symbol);
    if (!data || !data.price) {
      logError('updateSingleThaiPrice', 'ไม่พบราคา: ' + symbol);
      return false;
    }

      sheet.getRange(START_ROW.HOLD + i, HOLD_COL.PRICE_NOW).setValue(data.price);
    logInfo('updateSingleThaiPrice', 'อัปเดต ' + ticker + ' = ' + data.price);
    _clearHoldingsDataMemo();   // ✅ เขียน PRICE_NOW แล้ว เคลียร์ memo กัน getHoldingsData() คืนค่าเก่า
    return true;
  }

  return false; // ไม่พบ ticker ในชีต (ไม่ควรเกิด เพราะ processBuyAndCheckStockMode สร้างแถวไว้ก่อนแล้ว)
}


function updateSingleUSPrice(ticker) {
  ticker = String(ticker || '').trim().toUpperCase();
  if (!ticker) return false;

  const sheet   = getSheet(SHEETS.US_HOLD);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.HOLD) return false;

  const numRows = lastRow - START_ROW.HOLD + 1;
  const rows = sheet.getRange(START_ROW.HOLD, 1, numRows, 15).getValues();

  for (let i = 0; i < rows.length; i++) {
    const rowTicker = String(rows[i][HOLD_COL.TICKER - 1] || '').trim().toUpperCase();
    if (rowTicker !== ticker) continue;

    const sharesRemain = Number(rows[i][HOLD_COL.SHARES_REMAIN - 1]) || 0;
    if (sharesRemain <= 0) return false;

    const data = fetchYahooPrice(ticker);
    if (!data || !data.price) {
      logError('updateSingleUSPrice', 'ไม่พบราคา: ' + ticker);
      return false;
    }

     sheet.getRange(START_ROW.HOLD + i, HOLD_COL.PRICE_NOW).setValue(data.price);
    logInfo('updateSingleUSPrice', 'อัปเดต ' + ticker + ' = ' + data.price);
    _clearHoldingsDataMemo();   // ✅ เขียน PRICE_NOW แล้ว เคลียร์ memo กัน getHoldingsData() คืนค่าเก่า
    return true;
  }

  return false;
}




// ----------------------------------------
// /update — Update Setting รวม
// ----------------------------------------
function updateSetting() {
  try {
    sendTelegramSafe("⏳ กำลัง Update Setting...");

    updateThaiPrice();
    Utilities.sleep(1000);
    updateUSPrice();       // ✅ เพิ่ม US
    Utilities.sleep(1000);
    // ❌ ตัด updateThaiHigh() ออก
    resetAlertStatusAuto();
    resetUSDTHBStatusAuto();

    sendTelegramSafe(
      "✅ Update Setting สำเร็จ\n\n" +
      "🇹🇭 อัปเดตราคาหุ้นไทยแล้ว\n" +
      "🇺🇸 อัปเดตราคาหุ้นสหรัฐแล้ว\n" +
      "🔄 Reset Alert Status แล้ว\n" +
      "💱 Reset FX Status แล้ว\n\n" +
      "🕐 " + getNow()
    );
  } catch (e) {
    sendTelegramError("updateSetting", e);
  }
}






// ----------------------------------------
// Trigger
// ----------------------------------------
function createUpdateTrigger() {
  ["updateThaiPrice", "updateUSPriceTrigger"]  // ❌ ตัด updateThaiHigh ออก
    .forEach(fnName => {
      ScriptApp.getProjectTriggers()
        .filter(t => t.getHandlerFunction() === fnName)
        .forEach(t => ScriptApp.deleteTrigger(t));
    });

  ScriptApp.newTrigger("updateThaiPrice")
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger("updateUSPriceTrigger")
    .timeBased().everyMinutes(15).create();

  Logger.log("✅ Update Triggers: TH+US Price ทุก 15 นาที");
}


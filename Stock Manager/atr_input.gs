// ========================================
// atr_input.gs
// /atr14 — บันทึกค่า ATR 14 ลง Backend_ATR
// /atr14list — ดูรายการ ATR ทั้งหมด
// ========================================

// ----------------------------------------
// /atr14 TICKER VALUE
// ตัวอย่าง:
//   /atr14 VOO 9.15
//   /atr14 SCB 2.20
// ----------------------------------------
function cmdUpdateATR14(text) {
  try {
    const parts = text.trim().split(/\s+/);

    if (parts.length < 3) {
      sendTelegramSafe(
        "❌ รูปแบบไม่ถูกต้องครับ\n\n" +
        "✅ วิธีใช้:\n" +
        "/atr14 TICKER ราคา\n\n" +
        "📌 ตัวอย่าง:\n" +
        "/atr14 VOO 9.15\n" +
        "/atr14 SCB 2.20\n\n" +
        "💡 พิมพ์ /atr14list เพื่อดูรายการ"
      );
      return;
    }

    const ticker = parts[1].toUpperCase().trim();
    const atrVal = parseFloat(parts[2]);

    if (isNaN(atrVal) || atrVal < 0) {
      sendTelegramSafe(
        "❌ ค่า ATR ไม่ถูกต้อง: " + parts[2] + "\n" +
        "กรุณาใส่ตัวเลขทศนิยม เช่น 9.15"
      );
      return;
    }

    const result = _writeATR14(ticker, atrVal);

    if (result.success) {
      sendTelegramSafe(
        "✅ บันทึก ATR 14 สำเร็จ\n\n" +
        "📌 Ticker  : " + result.ticker + " (" + result.market + ")\n" +
        "📊 ATR ใหม่ : " + fmt(atrVal) + "\n" +
        "📊 ATR เดิม : " + fmt(result.oldATR) + "\n" +
        "📝 แถวที่   : " + result.row + "\n" +
        "🕐 " + getNow()
      );
    } else {
      sendTelegramSafe(
        "❌ ไม่พบ Ticker: \"" + ticker + "\"\n\n" +
        "💡 พิมพ์ /atr14list เพื่อดูรายการ Ticker ที่มี\n" +
        "หรือเพิ่ม Ticker ใน Sheet Backend_ATR ก่อนครับ"
      );
    }

  } catch (e) {
    sendTelegramError("cmdUpdateATR14", e);
  }
}

// ----------------------------------------
// /atr14list — แสดงรายการ ATR ทั้งหมด
// ----------------------------------------
function sendATR14List() {
  try {
    const sheet   = getSheet(BACKEND_ATR.SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow < BACKEND_ATR.START_ROW) {
      sendTelegramSafe("📭 ไม่มีข้อมูล ATR ในระบบ");
      return;
    }

    const numRows = lastRow - BACKEND_ATR.START_ROW + 1;
    const rows    = sheet.getRange(
      BACKEND_ATR.START_ROW, 1, numRows, 6
    ).getValues();

    let usMsg = "";
    let thMsg = "";

    rows.forEach(row => {
      const usTicker = String(row[BACKEND_ATR.US_TICKER - 1] || "").trim();
      const usATR    = Number(row[BACKEND_ATR.US_ATR    - 1]) || 0;
      const thTicker = String(row[BACKEND_ATR.TH_TICKER - 1] || "").trim();
      const thATR    = Number(row[BACKEND_ATR.TH_ATR    - 1]) || 0;

      if (usTicker) {
        usMsg += "  • " + usTicker.padEnd(8) + " : " + fmt(usATR) + "\n";
      }
      if (thTicker) {
        thMsg += "  • " + thTicker.padEnd(8) + " : " + fmt(thATR) + "\n";
      }
    });

    let msg =
      "📊 ATR 14 — รายการทั้งหมด\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "🇺🇸 หุ้นสหรัฐ\n" +
      (usMsg || "  — ไม่มีข้อมูล\n") + "\n" +
      "━━━━━━━━━━━━\n" +
      "🇹🇭 หุ้นไทย\n" +
      (thMsg || "  — ไม่มีข้อมูล\n") + "\n" +
      "━━━━━━━━━━━━\n" +
      "💡 อัปเดต ATR:\n" +
      "/atr14 TICKER ราคา\n" +
      "เช่น /atr14 VOO 9.15";

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendATR14List", e);
  }
}

// ----------------------------------------
// Private — เขียน ATR ลงชีท
// ค้นหาทั้ง US (col A) และ TH (col E)
// ----------------------------------------
function _writeATR14(ticker, atrVal) {
  const sheet   = getSheet(BACKEND_ATR.SHEET);

  // แก้ _writeATR14() ให้ใช้ BACKEND_ATR จาก config
  // BACKEND_ATR.TH_TICKER = 3 (col C)
  // BACKEND_ATR.TH_ATR    = 4 (col D)
 
  const lastRow = sheet.getLastRow();
  if (lastRow < BACKEND_ATR.START_ROW) return { success: false };

  const numRows = lastRow - BACKEND_ATR.START_ROW + 1;
  const rows    = sheet.getRange(
    BACKEND_ATR.START_ROW, 1, numRows, 6
  ).getValues();

  for (let i = 0; i < rows.length; i++) {
    const rowNum   = BACKEND_ATR.START_ROW + i;
    const usTicker = String(rows[i][BACKEND_ATR.US_TICKER - 1] || "").trim().toUpperCase();
    const thTicker = String(rows[i][BACKEND_ATR.TH_TICKER - 1] || "").trim().toUpperCase();

    // เจอใน US (col A/B)
    if (usTicker === ticker) {
      const oldATR = Number(rows[i][BACKEND_ATR.US_ATR - 1]) || 0;
      sheet.getRange(rowNum, BACKEND_ATR.US_ATR).setValue(atrVal);
      return { success: true, ticker, market: "🇺🇸 US", oldATR, row: rowNum };
    }

    // เจอใน TH (col E/F)
    if (thTicker === ticker) {
      const oldATR = Number(rows[i][BACKEND_ATR.TH_ATR - 1]) || 0;
      sheet.getRange(rowNum, BACKEND_ATR.TH_ATR).setValue(atrVal);
      return { success: true, ticker, market: "🇹🇭 TH", oldATR, row: rowNum };
    }
  }

  return { success: false };
}


// ----------------------------------------
// Auto Trigger — แจ้งเตือนกรอก ATR 14
// ทุกวันอาทิตย์ 19:00
// ----------------------------------------
function sendATR14Reminder() {
  try {
    const sheet   = getSheet(BACKEND_ATR.SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow < BACKEND_ATR.START_ROW) return;

    const numRows = lastRow - BACKEND_ATR.START_ROW + 1;
    const rows    = sheet.getRange(
      BACKEND_ATR.START_ROW, 1, numRows, 6
    ).getValues();

    let usList = [];
    let thList = [];

    rows.forEach(row => {
      const usTicker = String(row[BACKEND_ATR.US_TICKER - 1] || "").trim();
      const usATR    = Number(row[BACKEND_ATR.US_ATR    - 1]) || 0;
      const thTicker = String(row[BACKEND_ATR.TH_TICKER - 1] || "").trim();
      const thATR    = Number(row[BACKEND_ATR.TH_ATR    - 1]) || 0;

      if (usTicker) usList.push({ ticker: usTicker, atr: usATR });
      if (thTicker) thList.push({ ticker: thTicker, atr: thATR });
    });

    if (usList.length === 0 && thList.length === 0) return;

    let msg =
      "🔔 แจ้งเตือน — กรอก ATR 14 ประจำสัปดาห์\n" +
      "📅 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";

    if (usList.length > 0) {
      msg += "🇺🇸 หุ้นสหรัฐ\n";
      usList.forEach((h, i) => {
        msg += (i + 1) + ". " + h.ticker +
          " | ATR ล่าสุด : " + fmt(h.atr) + "\n";
      });
      msg += "\n";
    }

    if (thList.length > 0) {
      msg += "🇹🇭 หุ้นไทย\n";
      thList.forEach((h, i) => {
        msg += (i + 1) + ". " + h.ticker +
          " | ATR ล่าสุด : " + fmt(h.atr) + "\n";
      });
      msg += "\n";
    }

    msg +=
      "━━━━━━━━━━━━\n" +
      "💡 วิธีกรอก:\n" +
      "/atr14 TICKER ค่า\n\n" +
      "📌 ตัวอย่าง:\n" +
      "/atr14 VOO 9.15\n" +
      "/atr14 SCB 2.20";

    sendTelegramSafe(msg);
  } catch (e) {
    logError("sendATR14Reminder", e);
  }
}

// ----------------------------------------
// Setup Trigger ATR14 Reminder
// ทุกวันอาทิตย์ 19:00
// ----------------------------------------
function createATR14ReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendATR14Reminder")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendATR14Reminder")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(19)
    .inTimezone("Asia/Bangkok")
    .create();

  Logger.log("✅ ATR14 Reminder Trigger: ทุกวันอาทิตย์ 19:00");
}



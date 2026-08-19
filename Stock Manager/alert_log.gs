// ========================================
// alert_log.gs
// บันทึก Alert History ทุกครั้งที่ส่ง
// ========================================

function _logAlert(type, symbol, message) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName("Alert_History");

    // สร้าง sheet ถ้ายังไม่มี
    if (!sheet) {
      sheet = ss.insertSheet("Alert_History");
      sheet.getRange(1, 1, 1, 5).setValues([[
        "Timestamp", "Type", "Symbol", "Message", "Status"
      ]]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      type,
      symbol || "",
      message.substring(0, 500), // จำกัด 500 ตัวอักษร
      "SENT"
    ]);

  } catch (e) {
    logError("_logAlert", e);
  }
}

// ----------------------------------------
// /alertlog — ดู Alert ล่าสุด 20 รายการ
// ----------------------------------------
function sendAlertLog() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Alert_History");

    if (!sheet || sheet.getLastRow() <= 1) {
      sendTelegramSafe("📭 ยังไม่มี Alert History");
      return;
    }

    const lastRow = sheet.getLastRow();
    const start   = Math.max(2, lastRow - 19); // 20 รายการล่าสุด
    const rows    = sheet.getRange(
      start, 1, lastRow - start + 1, 5
    ).getValues();

    let msg =
      "📋 Alert History (20 ล่าสุด)\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";

    rows.reverse().forEach(row => {
      const ts      = row[0] instanceof Date
        ? Utilities.formatDate(row[0], "Asia/Bangkok", "dd/MM HH:mm")
        : String(row[0]);
      const type   = row[1] || "";
      const symbol = row[2] || "";
      const status = row[4] || "";

      msg += "🕐 " + ts + "\n" +
             "  [" + type + "] " + (symbol ? symbol + " — " : "") +
             status + "\n\n";
    });

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendAlertLog", e);
  }
}

function createAlertHistorySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName("Alert_History");
  if (sheet) {
    Logger.log("มีอยู่แล้ว");
    return;
  }

  sheet = ss.insertSheet("Alert_History");
  sheet.getRange(1, 1, 1, 5).setValues([[
    "Timestamp", "Type", "Symbol", "Message", "Status"
  ]]);
  sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  sheet.setFrozenRows(1);
  Logger.log("✅ สร้าง Alert_History sheet แล้ว");
}





// ----------------------------------------
// /alertclear — ล้าง Alert History
// ----------------------------------------
function clearAlertLog() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Alert_History");
    if (!sheet) {
      sendTelegramSafe("📭 ไม่พบ Alert_History sheet");
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    sendTelegramSafe(
      "✅ ล้าง Alert History สำเร็จ\n🕐 " + getNow()
    );
  } catch (e) {
    sendTelegramError("clearAlertLog", e);
  }
}

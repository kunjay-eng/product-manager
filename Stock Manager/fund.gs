// ========================================
// fund.gs
// /nav — บันทึก NAV กองทุน
// /navlist — ดูรายการกองทุนและ NAV ล่าสุด
// ========================================

const FUND_HOLD_SHEET = "🏛️​ 💼 Holdings";
const FUND_START_ROW  = 4;
const FUND_COL_NAME   = 2;   // B ชื่อกองทุน
const FUND_COL_NAV    = 8;   // H NAV ปัจจุบัน

// ----------------------------------------
// /nav FUNDNAME VALUE
// ตัวอย่าง: /nav KKP-NDQ100-UH-E 17.6160
// ----------------------------------------
function cmdUpdateNAV(text) {
  try {
    // parse command
    // รูปแบบ: /nav <ชื่อกองทุน> <ราคา>
    const parts = text.trim().split(/\s+/);
    // parts[0] = /nav
    // parts[1] = fund name (อาจมีช่องว่าง)
    // parts[parts.length-1] = ราคา NAV

    if (parts.length < 3) {
      sendTelegramSafe(
        "❌ รูปแบบไม่ถูกต้องครับ\n\n" +
        "✅ วิธีใช้:\n" +
        "/nav ชื่อกองทุน ราคา\n\n" +
        "📌 ตัวอย่าง:\n" +
        "/nav KKP  NDQ100-UH-E 17.6160\n" +
        "หรือ กดค้าง แล้วใส่ราคา\n" +
        "/nav KKP  NDQ100-UH-E \n\n" +
        "💡 พิมพ์ /navlist เพื่อดูรายการกองทุน"
      );
      return;
    }

    // ดึงราคาจาก parts สุดท้าย
    const navStr  = parts[parts.length - 1];
    const navVal  = parseFloat(navStr);

    if (isNaN(navVal) || navVal <= 0) {
      sendTelegramSafe(
        "❌ ราคา NAV ไม่ถูกต้อง: " + navStr + "\n" +
        "กรุณาใส่ตัวเลขทศนิยม เช่น 17.6160"
      );
      return;
    }

    // ชื่อกองทุน = ทุกอย่างระหว่าง /nav กับ ราคา
    const fundName = parts.slice(1, parts.length - 1).join(" ").trim();

    // ค้นหาในชีท
    const result = _writeFundNAV(fundName, navVal);

    if (result.success) {
      sendTelegramSafe(
        "✅ บันทึก NAV สำเร็จ\n\n" +
        "🏛️ กองทุน  : " + result.matchedName + "\n" +
        "📈 NAV ใหม่ : " + fmtTHB(navVal) + "\n" +
        "📈 NAV เดิม : " + fmtTHB(result.oldNav) + "\n" +
        "📝 แถวที่   : " + result.row + "\n" +
        "🕐 " + getNow()
      );
    } else {
      // ไม่พบกองทุน — แสดงรายชื่อที่มี
      const list = _getFundList();
      let msg =
        "❌ ไม่พบกองทุน: \"" + fundName + "\"\n\n" +
        "📋 กองทุนที่มีในระบบ:\n";
      list.forEach((f, i) => {
        msg += "\nav " + f.name + "\n";
      });
      msg += "\n💡 คัดลอกชื่อจากรายการด้านบนมาใช้ได้เลยครับ";
      sendTelegramSafe(msg);
    }

  } catch (e) {
    sendTelegramError("cmdUpdateNAV", e);
  }
}

// ----------------------------------------
// /navlist — แสดงรายการกองทุนและ NAV ล่าสุด
// ----------------------------------------
function sendFundNAVList() {
  try {
    const list = _getFundList();

    if (list.length === 0) {
      sendTelegramSafe("📭 ไม่มีกองทุนในระบบ");
      return;
    }

    let msg =
      "🏛️ รายการกองทุนรวม\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";

    list.forEach((f, i) => {
      const plPct  = f.avgCost > 0 && f.navNow > 0
        ? ((f.navNow - f.avgCost) / f.avgCost) * 100 : 0;

      msg +=
        (i + 1) + ". " + f.name + "\n" +
        "  💰 avg cost : " + fmtTHB(f.avgCost)    + "\n" +
        "  📈 NAV      : " + fmtTHB(f.navNow)      + "\n" +
        "  📊 P&L      : " + signPct(plPct)         + "\n" +
        "  💵 มูลค่า   : " + fmtTHB(f.valueNow)    + "\n\n";
    });

    msg +=
      "━━━━━━━━━━━━\n" +
      "💡 อัปเดต NAV:\n" +
      "/nav ชื่อกองทุน ราคา\n" +
      "เช่น /nav KKP-NDQ100-UH-E 17.6160";

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendFundNAVList", e);
  }
}

// ----------------------------------------
// Private — เขียน NAV ลงชีท
// ----------------------------------------
function _writeFundNAV(fundName, navVal) {
  const sheet   = getSheet(FUND_HOLD_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow < FUND_START_ROW) {
    return { success: false };
  }

  const numRows = lastRow - FUND_START_ROW + 1;
  const rows    = sheet.getRange(
    FUND_START_ROW, 1, numRows, 12
  ).getValues();

  // ค้นหาชื่อกองทุน (case-insensitive, trim)
  const searchName = fundName.toLowerCase().replace(/\s+/g, "");

  for (let i = 0; i < rows.length; i++) {
    const name = String(rows[i][FUND_COL_NAME - 1] || "").trim();
    if (!name) continue;

    const compareName = name.toLowerCase().replace(/\s+/g, "");

    // เทียบแบบ exact หรือ includes
    if (compareName === searchName || compareName.includes(searchName) || searchName.includes(compareName)) {
      const rowNum = FUND_START_ROW + i;
      const oldNav = Number(rows[i][FUND_COL_NAV - 1]) || 0;

      // เขียน NAV ลง col H
      sheet.getRange(rowNum, FUND_COL_NAV).setValue(navVal);

      return {
        success:     true,
        matchedName: name,
        oldNav:      oldNav,
        row:         rowNum
      };
    }
  }

  return { success: false };
}

// ----------------------------------------
// Private — ดึงรายชื่อกองทุนทั้งหมด
// ----------------------------------------
function _getFundList() {
  const sheet   = getSheet(FUND_HOLD_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow < FUND_START_ROW) return [];

  const numRows = lastRow - FUND_START_ROW + 1;
  const rows    = sheet.getRange(
    FUND_START_ROW, 1, numRows, 12
  ).getValues();

  const result = [];

  for (const row of rows) {
    const name        = String(row[FUND_COL_NAME - 1] || "").trim();
    const unitsRemain = Number(row[FUND_HOLD_COL.UNITS_REMAIN - 1]) || 0;
    if (!name || unitsRemain <= 0) continue;

    result.push({
      name,
      unitsRemain,
      avgCost:  Number(row[FUND_HOLD_COL.AVG_COST   - 1]) || 0,
      navNow:   Number(row[FUND_HOLD_COL.NAV_NOW    - 1]) || 0,
      valueNow: Number(row[FUND_HOLD_COL.VALUE_NOW  - 1]) || 0,
    });
  }

  return result;
}


// ----------------------------------------
// Auto Trigger — แจ้งเตือนกรอก NAV
// จันทร์-เสาร์ 18:00
// ----------------------------------------
function sendNAVReminder() {
  try {
    // เช็คว่าวันนี้ไม่ใช่วันอาทิตย์
    const today = new Date();
    const day   = today.getDay(); // 0=อาทิตย์, 1=จันทร์, ..., 6=เสาร์
    if (day === 0) return;        // วันอาทิตย์ไม่ส่ง

    const list = _getFundList();
    if (list.length === 0) return;

    let msg =
      "🔔 แจ้งเตือน — กรอก NAV กองทุน\n" +
      "📅 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "📋 กองทุนที่ต้องอัปเดต NAV:\n\n";

    list.forEach((f, i) => {
      msg +=
        (i + 1) + ". " + f.name + "\n" +
        "   NAV ล่าสุด : " + fmtTHB(f.navNow) + "\n\n";
    });

    msg +=
      "━━━━━━━━━━━━\n" +
      "💡 วิธีกรอก:\n" +
      "/nav ชื่อกองทุน ราคา\n\n" +
      "📌 ตัวอย่าง:\n" +
      "/nav KKP-NDQ100-UH-E 17.6160";

    sendTelegramSafe(msg);
  } catch (e) {
    logError("sendNAVReminder", e);
  }
}

// ----------------------------------------
// Setup Trigger NAV Reminder
// จันทร์-เสาร์ 18:00 (ตั้ง trigger ทุกวัน
// แล้วเช็ควันใน function)
// ----------------------------------------
function createNAVReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendNAVReminder")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendNAVReminder")
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .inTimezone("Asia/Bangkok")
    .create();

  Logger.log("✅ NAV Reminder Trigger: ทุกวัน 18:00 (ข้ามวันอาทิตย์อัตโนมัติ)");
}


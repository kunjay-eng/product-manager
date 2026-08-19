// ========================================
// portfolio.gs
// /dca — DCA Trigger Alert
// ========================================

const MONTH_TH = [
  "","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม",
  "มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม",
  "พฤศจิกายน","ธันวาคม"
];

// ----------------------------------------
// Status Helper
// ----------------------------------------
function _isTriggered(status) {
  return status.includes("✅") && status.includes("Trigger");
}

function _isNearTrigger(status) {
  return status.includes("⚠️") || status.includes("ใกล้");
}

// ----------------------------------------
// /dca — Command สั่งดูสถานะทั้งหมด
// ----------------------------------------
function sendPortfolioDCA() {
  sendTelegramSafe("⏳ กำลังตรวจสอบ DCA Trigger...");
  try {
    _sendDCAReport(SHEETS.TH_PORT, "🇹🇭 หุ้นไทย",   "฿", true);
    _sendDCAReport(SHEETS.US_PORT, "🇺🇸 หุ้นสหรัฐ", "$", true);
  } catch (e) {
    sendTelegramError("sendPortfolioDCA", e);
  }
}

// ----------------------------------------
// Auto Trigger — ทุก 30 นาที
// ส่งเฉพาะเมื่อ Triggered จริง
// ----------------------------------------
function checkPortfolioTriggerAuto() {
  try {
    _sendDCAReport(SHEETS.TH_PORT, "🇹🇭 หุ้นไทย",   "฿", false);
    _sendDCAReport(SHEETS.US_PORT, "🇺🇸 หุ้นสหรัฐ", "$", false);
  } catch (e) {
    logError("checkPortfolioTriggerAuto", e);
  }
}

// ----------------------------------------
// Private — Build + Send DCA Report
// ----------------------------------------
function _sendDCAReport(sheetName, label, currency, forceAll) {
  const rows    = getPortfolioRows(sheetName);
  const fxRate  = getFxRate();
  const isTH    = currency === "฿";

  // แยกกลุ่ม
  const triggered   = rows.filter(r => _isTriggered(r.status));
  const nearTrigger = rows.filter(r => _isNearTrigger(r.status));
  const withCond    = rows.filter(r => r.condCost > 0 || r.condPrice > 0);



  // Auto mode — ส่งเฉพาะ triggered
  if (!forceAll) {
    if (triggered.length === 0) return;

    let msg =
      "🎯 DCA TRIGGER ALERT!\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n" +
      label + "\n\n";

    for (const r of triggered) {
      
  const costLabel = r.costReduce >= 0 ? "✂️ ลดต้นทุน" : "📈 เพิ่มต้นทุน";
  const costValue = Math.abs(r.costReduce);



      const budgetShares = r.budget > 0 && r.trigger > 0
        ? (isTH
          ? r.budget / r.trigger
          : (r.budget / fxRate) / r.trigger)
        : 0;

     msg +=
  "✅ " + r.ticker + " ถึง Trigger!\n" +
  "  💰 ต้นทุน avg    : " + currency + fmt(r.buyPrice)    + "\n" +
  "  📈 ราคาปัจจุบัน : " + currency + fmt(r.priceNow)    + "\n" +
  "  🎯 ราคา Trigger  : " + currency + fmt(r.trigger)     + "\n" +
  "  💵 งบประมาณ      : " + fmtTHB(r.budget)             + "\n" +
  "  📦 ซื้อได้ประมาณ : " + fmt(budgetShares) + " หุ้น"  + "\n" +
  "  📊 avg ใหม่       : " + currency + fmt(r.newAvgCost) + "\n" +
  "  " + costLabel + "     : " + currency + fmt(costValue) + "\n\n";

    }

    sendTelegramSafe(msg);
    return;
  }

  // Force mode — แสดงทุก row ที่มีเงื่อนไข
  if (withCond.length === 0) {
    sendTelegramSafe(
      label + "\n" +
      "📭 ยังไม่มีหุ้นที่ตั้งเงื่อนไขซื้อเพิ่ม"
    );
    return;
  }

  let msg =
    "📊 DCA STATUS REPORT\n" +
    "🕐 " + getNow() + "\n" +
    "━━━━━━━━━━━━\n" +
    label + " (" + withCond.length + " ตัว)\n\n";

  // Triggered
  if (triggered.length > 0) {
    msg += "✅ ถึง Trigger แล้ว (" + triggered.length + " ตัว)\n";
    for (const r of triggered) {
      const budgetShares = r.budget > 0 && r.trigger > 0
        ? (isTH
          ? r.budget / r.trigger
          : (r.budget / fxRate) / r.trigger)
        : 0;

const costLabel = r.costReduce >= 0 ? "✂️ ลดต้นทุน" : "📈 เพิ่มต้นทุน";
const costValue = Math.abs(r.costReduce);
  

      msg +=
  "━━━━━━━━━━━━\n" +
  "✅ " + r.ticker + "\n" +
  "  💰 avg          : " + currency + fmt(r.buyPrice)    + "\n" +
  "  📈 ปัจจุบัน    : " + currency + fmt(r.priceNow)    + "\n" +
  "  🎯 Trigger      : " + currency + fmt(r.trigger)     + "\n" +
  "  💵 งบ           : " + fmtTHB(r.budget)             + "\n" +
  "  📦 ซื้อได้       : " + fmt(budgetShares) + " หุ้น"  + "\n" +
  "  📊 avg ใหม่      : " + currency + fmt(r.newAvgCost) + "\n" +
  "  " + costLabel + "     : " + currency + fmt(costValue) + "\n\n";

    }
  }

  // Near Trigger
  if (nearTrigger.length > 0) {
    msg += "⚠️ ใกล้ Trigger (" + nearTrigger.length + " ตัว)\n";
    for (const r of nearTrigger) {
      const diff = r.priceNow > 0 && r.trigger > 0
        ? ((r.priceNow - r.trigger) / r.trigger) * 100
        : 0;

      msg +=
        "━━━━━━━━━━━━\n" +
        "⚠️ " + r.ticker + "\n" +
        "  💰 avg       : " + currency + fmt(r.buyPrice)  + "\n" +
        "  📈 ปัจจุบัน : " + currency + fmt(r.priceNow)  + "\n" +
        "  🎯 Trigger   : " + currency + fmt(r.trigger)   + "\n" +
        "  📏 ห่าง      : " + signPct(diff)               + "\n" +
        "  💵 งบ        : " + fmtTHB(r.budget)            + "\n\n";
    }
  }

  // Safe — ยังไม่ถึง
  const safe = withCond.filter(
    r => !_isTriggered(r.status) && !_isNearTrigger(r.status)
  );
  if (safe.length > 0) {
    msg += "🔵 ยังไม่ถึงเงื่อนไข (" + safe.length + " ตัว)\n\n";
    for (const r of safe) {
      const diff = r.priceNow > 0 && r.trigger > 0
        ? ((r.priceNow - r.trigger) / r.trigger) * 100
        : 0;

      msg +=
        "  • " + r.ticker +
        " | ปัจจุบัน " + currency + fmt(r.priceNow) + "\n" +
        " | ราคา trigger " + currency + fmt(r.trigger) +
        " | ห่าง " + signPct(diff) + "\n\n";
    }
  }

  sendTelegramSafe(msg);
}

// ----------------------------------------
// Trigger Setup
// ----------------------------------------
function createPortfolioTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "checkPortfolioTriggerAuto")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("checkPortfolioTriggerAuto")
    .timeBased().everyMinutes(30).create();

  Logger.log("✅ Portfolio Trigger: ทุก 30 นาที");
}

// ----------------------------------------
// /allstocks — แสดงหุ้นทั้งหมดในพอร์ต (ไม่กรองเงื่อนไข)
// ----------------------------------------
function sendAllStocks() {
  sendTelegramSafe("⏳ กำลังดึงข้อมูลหุ้นทั้งหมด...");
  try {
    _sendAllStocksReport(SHEETS.TH_PORT, "🇹🇭 หุ้นไทย",   "฿");
    _sendAllStocksReport(SHEETS.US_PORT, "🇺🇸 หุ้นสหรัฐ", "$");
  } catch (e) {
    sendTelegramError("sendAllStocks", e);
  }
}


// ----------------------------------------
// Private — Build All Stocks Report
// ----------------------------------------
function _sendAllStocksReport(sheetName, label, currency) {
  const rows = getPortfolioRows(sheetName);

  if (rows.length === 0) {
    sendTelegramSafe(label + "\n📭 ไม่มีหุ้นในพอร์ต");
    return;
  }

  let msg =
    "📋 ALL STOCKS\n" +
    "🕐 " + getNow() + "\n" +
    "━━━━━━━━━━━━\n" +
    label + " (" + rows.length + " ตัว)\n\n";

  for (const r of rows) {
    const plPct = r.buyPrice > 0 && r.priceNow > 0
      ? ((r.priceNow - r.buyPrice) / r.buyPrice) * 100
      : 0;

    const hasCond = r.condCost > 0 || r.condPrice > 0;

    msg +=
      plEmoji(plPct) + " " + r.ticker + "\n" +
      "  💰 avg     : " + currency + fmt(r.buyPrice) + "\n" +
      "  📈 ปัจจุบัน : " + currency + fmt(r.priceNow) + "\n" +
      "  📊 P/L     : " + signPct(plPct) + "\n";

    if (hasCond) {
      msg +=
        "  🎯 Trigger : " + currency + fmt(r.trigger) +
        "  [" + r.status + "]\n";
    } else {
      msg += "  🎯 Trigger : -- (ไม่ได้ตั้งเงื่อนไข)\n";
    }

    msg += "\n";
  }

  sendTelegramSafe(msg);
}

// ----------------------------------------
// Setup Triggers
// ----------------------------------------

function createPortfolioTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "checkPortfolioTriggerAuto")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // #7 Portfolio DCA Check — ทุก 30 นาที
  ScriptApp.newTrigger("checkPortfolioTriggerAuto")
    .timeBased().everyMinutes(30).create();

  Logger.log("✅ Portfolio DCA Trigger: ทุก 30 นาที");
}




// ========================================
// fx.gs
// ========================================

// ----------------------------------------
// ดึงข้อมูล FX ทั้งหมดจาก Sheet
// ----------------------------------------
function _getFXData() {
  const sheet = getSheet(SHEETS.FX);
  return {
    rate:       Number(sheet.getRange(FX_CELL.RATE).getValue())       || 0,
    buy1:       Number(sheet.getRange(FX_CELL.BUY1).getValue())       || 0,
    buy2:       Number(sheet.getRange(FX_CELL.BUY2).getValue())       || 0,
    sell:       Number(sheet.getRange(FX_CELL.SELL).getValue())       || 0,
    flagBuy1:   String(sheet.getRange(FX_CELL.FLAG_BUY1).getValue())  || "WAITING",
    flagBuy2:   String(sheet.getRange(FX_CELL.FLAG_BUY2).getValue())  || "WAITING",
    flagSell:   String(sheet.getRange(FX_CELL.FLAG_SELL).getValue())  || "WAITING",
    lastCheck:  sheet.getRange(FX_CELL.LAST_CHECK).getValue()
  };
}

// ----------------------------------------
// Auto Trigger — ตรวจ FX Zone ทุก 15 นาที
// ----------------------------------------
function checkUSDTHBAuto() {
  try {
    const sheet = getSheet(SHEETS.FX);
    const d     = _getFXData();

    if (d.rate <= 0) return;

    const now    = getNow();
    let   alerted = false;

    // เช็ค Buy1 (rate <= buy1)
    if (d.rate <= d.buy1 && d.flagBuy1 === "WAITING") {
      sendTelegramSafe(
        "💱 FX ALERT — Buy Zone 1\n\n" +
        "🇺🇸 USD/THB : " + fmt4(d.rate) + "\n" +
        "📉 แตะ Buy Zone 1 ≤ " + fmt4(d.buy1) + "\n\n" +
        "💡 จังหวะแลกเงินดี — Zone 1\n" +
        "🕐 " + now
      );
      sheet.getRange(FX_CELL.FLAG_BUY1).setValue("SENT");
      alerted = true;
    }

    // เช็ค Buy2 (rate <= buy2)
    if (d.rate <= d.buy2 && d.flagBuy2 === "WAITING") {
      sendTelegramSafe(
        "💱 FX ALERT — Buy Zone 2 🔥\n\n" +
        "🇺🇸 USD/THB : " + fmt4(d.rate) + "\n" +
        "📉 แตะ Buy Zone 2 ≤ " + fmt4(d.buy2) + "\n\n" +
        "💡 จังหวะแลกเงินดีมาก — Zone 2\n" +
        "🕐 " + now
      );
      sheet.getRange(FX_CELL.FLAG_BUY2).setValue("SENT");
      alerted = true;
    }

    // เช็ค Sell (rate >= sell)
    if (d.rate >= d.sell && d.flagSell === "WAITING") {
      sendTelegramSafe(
        "💱 FX ALERT — Sell Zone\n\n" +
        "🇺🇸 USD/THB : " + fmt4(d.rate) + "\n" +
        "📈 แตะ Sell Zone ≥ " + fmt4(d.sell) + "\n\n" +
        "💡 จังหวะขายดอลล่าร์/โอนกลับ\n" +
        "🕐 " + now
      );
      sheet.getRange(FX_CELL.FLAG_SELL).setValue("SENT");
      alerted = true;
    }

    // อัปเดต Last_Check เสมอ
    sheet.getRange(FX_CELL.LAST_CHECK).setValue(new Date());

  } catch (e) {
    logError("checkUSDTHBAuto", e);
  }
}


// ----------------------------------------
// /exchange จำนวน [thb|usd]
// ----------------------------------------
function cmdExchange(text) {
  try {
    const parts   = text.trim().split(/\s+/);
    const amount  = parseFloat(parts[1]);
    const fromCur = (parts[2] || "thb").toLowerCase();

    if (!amount || isNaN(amount) || amount <= 0) {
      sendTelegramSafe(
        "❌ รูปแบบไม่ถูกต้อง\n\n" +
        "✅ วิธีใช้:\n" +
        "/exchange จำนวน [thb|usd]\n\n" +
        "📌 ตัวอย่าง:\n" +
        "/exchange 3000 thb  → บาท → USD\n" +
        "/exchange 100 usd   → USD → บาท\n" +
        "/exchange 3000      → บาท → USD (default)"
      );
      return;
    }

    if (fromCur !== "thb" && fromCur !== "usd") {
      sendTelegramSafe("❌ สกุลเงินต้องเป็น thb หรือ usd เท่านั้น");
      return;
    }

    const fxRate = getFxRate();

    let msg;
    if (fromCur === "thb") {
      const usd = amount / fxRate;
      msg =
        "💱 คำนวณแลกเงิน\n\n" +
        "🇹🇭 " + fmtTHB(amount) + "\n" +
        "         ↓\n" +
        "🇺🇸 " + fmtUSD(usd) + "\n\n" +
        "💹 อัตราแลกเปลี่ยน : " + fmt4(fxRate) + " THB/USD";
    } else {
      const thb = amount * fxRate;
      msg =
        "💱 คำนวณแลกเงิน\n\n" +
        "🇺🇸 " + fmtUSD(amount) + "\n" +
        "         ↓\n" +
        "🇹🇭 " + fmtTHB(thb) + "\n\n" +
        "💹 อัตราแลกเปลี่ยน : " + fmt4(fxRate) + " THB/USD";
    }

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("cmdExchange", e);
  }
}


// ----------------------------------------
// /checkfx — ดูสถานะ FX ปัจจุบัน
// ----------------------------------------
function sendFX() {
  try {
    const d = _getFXData();

    const rateEmoji =
      d.rate <= d.buy2 ? "🟢🟢" :
      d.rate <= d.buy1 ? "🟢"   :
      d.rate >= d.sell ? "🔴"   : "🟡";

    const zone =
      d.rate <= d.buy2 ? "Buy Zone 2 🔥 (ดีมาก)"  :
      d.rate <= d.buy1 ? "Buy Zone 1 (ดี)"         :
      d.rate >= d.sell ? "Sell Zone (ขายได้)"      : "Neutral Zone";

    // ✅ แก้ Last Check format เป็น ค.ศ.
    let lastCheckStr = "—";
    if (d.lastCheck instanceof Date) {
      lastCheckStr = Utilities.formatDate(
        d.lastCheck, "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss"
      );
    } else if (d.lastCheck) {
      lastCheckStr = String(d.lastCheck);
    }

    const msg =
      "💱 FX STATUS — USD/THB\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +

      rateEmoji + " ราคาปัจจุบัน : " + fmt4(d.rate) + "\n\n" +

      "📊 Alert Zones\n" +
      "  🟢🟢 Buy Zone 2 : ≤ " + fmt4(d.buy2) +
        (d.rate <= d.buy2 ? " ← HERE" : "") + "\n" +
      "  🟢   Buy Zone 1 : ≤ " + fmt4(d.buy1) +
        (d.rate <= d.buy1 && d.rate > d.buy2 ? " ← HERE" : "") + "\n" +
      "  🔴   Sell Zone  : ≥ " + fmt4(d.sell) +
        (d.rate >= d.sell ? " ← HERE" : "") + "\n\n" +

      "🎯 Zone ปัจจุบัน : " + zone + "\n\n" +

      "📋 Alert Status\n" +
      "  Buy1 Flag : " + (d.flagBuy1 === "SENT" ? "✅ SENT" : "⏳ WAITING") + "\n" +
      "  Buy2 Flag : " + (d.flagBuy2 === "SENT" ? "✅ SENT" : "⏳ WAITING") + "\n" +
      "  Sell Flag : " + (d.flagSell === "SENT" ? "✅ SENT" : "⏳ WAITING") + "\n\n" +

      "🕐 Last Check : " + lastCheckStr;  // ✅ ค.ศ. แล้ว

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendFX", e);
  }
}

// ----------------------------------------
// Reset FX — Command (ส่ง Telegram)
// ----------------------------------------
function resetUSDTHBStatus() {
  try {
    const sheet = getSheet(SHEETS.FX);
    sheet.getRange(FX_CELL.FLAG_BUY1).setValue("WAITING");
    sheet.getRange(FX_CELL.FLAG_BUY2).setValue("WAITING");
    sheet.getRange(FX_CELL.FLAG_SELL).setValue("WAITING");
    sheet.getRange(FX_CELL.LAST_CHECK).setValue(new Date());

    sendTelegramSafe(
      "✅ Reset FX Alert สำเร็จ\n\n" +
      "Buy1 Flag : ⏳ WAITING\n" +
      "Buy2 Flag : ⏳ WAITING\n" +
      "Sell Flag : ⏳ WAITING\n\n" +
      "🕐 " + getNow()
    );
  } catch (e) {
    sendTelegramError("resetUSDTHBStatus", e);
  }
}

// ----------------------------------------
// Reset FX — Auto (เงียบ)
// ----------------------------------------
function resetUSDTHBStatusAuto() {
  try {
    const sheet = getSheet(SHEETS.FX);
    sheet.getRange(FX_CELL.FLAG_BUY1).setValue("WAITING");
    sheet.getRange(FX_CELL.FLAG_BUY2).setValue("WAITING");
    sheet.getRange(FX_CELL.FLAG_SELL).setValue("WAITING");
    sheet.getRange(FX_CELL.LAST_CHECK).setValue(new Date());
    logInfo("resetUSDTHBStatusAuto", "Reset FX flags เรียบร้อย");
  } catch (e) {
    logError("resetUSDTHBStatusAuto", e);
  }
}

// ----------------------------------------
// Setup Trigger
// ----------------------------------------
function createFXTrigger() {
  ["checkUSDTHBAuto", "resetUSDTHBStatusAuto"].forEach(fnName => {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === fnName)
      .forEach(t => ScriptApp.deleteTrigger(t));
  });

  // เช็ค FX ทุก 15 นาที
  ScriptApp.newTrigger("checkUSDTHBAuto")
    .timeBased().everyMinutes(15).create();

  // ✅ Reset flags ทุกวัน 09:00 อัตโนมัติ
  ScriptApp.newTrigger("resetUSDTHBStatusAuto")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone("Asia/Bangkok")
    .create();

  Logger.log("✅ FX Trigger: check 15 นาที, reset ทุกวัน 09:00");
}

// ========================================
// holdings.gs
// /summary /dashboard /ranking /weekly*** /holdings***
// ========================================

/**
 * ====================================================
 * ฟังก์ชั่นนี้ไม่ใช้
 * ====================================================
 * 

// ----------------------------------------
// /summary
// ----------------------------------------
function sendPortfolioSummary() {
  sendTelegramSafe("⏳ กำลังสร้าง Portfolio Summary...");
  try {
    const d = collectAllData();

    const msg =
      "📊 PORTFOLIO SUMMARY\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "💱 USD/THB : " + fmt4(d.fxRate) + "\n\n" +

      "💵 Cash THB  : " + fmtTHB(d.cashTHB)   + "\n" +
      "💵 Cash USD  : " + fmtUSD(d.cashUSD)   + "\n" +
      "💵 Cash รวม  : " + fmtTHB(d.cashTotal) + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇹🇭 พอร์ตไทย\n" +
      "📈 มูลค่า      : " + fmtTHB(d.thPortTHB)   + "\n" +
      plEmoji(d.thUnrealTHB) +
        " Unrealized : " + signTHB(d.thUnrealTHB) + "\n" +
      plEmoji(d.thRealTHB) +
        " Realized   : " + signTHB(d.thRealTHB)   + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇺🇸 พอร์ตสหรัฐ\n" +
      "📈 มูลค่า      : " + fmtUSD(d.usPortUSD) +
        " (" + fmtTHB(d.usPortTHB) + ")\n" +
      plEmoji(d.usUnrealUSD) +
        " Unrealized : " + signUSD(d.usUnrealUSD) + "\n" +
      plEmoji(d.usRealUSD) +
        " Realized   : " + signUSD(d.usRealUSD)   + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🏛️ กองทุนรวม (" + d.fundHoldings.length + " กองทุน)\n" +
      "📈 มูลค่า      : " + fmtTHB(d.fundPortTHB)    + "\n" +
      plEmoji(d.fundUnrealTHB) +
        " Unrealized : " + signTHB(d.fundUnrealTHB)  + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🌎 พอร์ตรวม\n" +
      "📈 มูลค่า      : " + fmtTHB(d.totalPortTHB)    + "\n" +
      plEmoji(d.totalUnrealTHB) +
        " Unrealized : " + signTHB(d.totalUnrealTHB) + "\n" +
      plEmoji(d.totalRealTHB) +
        " Realized   : " + signTHB(d.totalRealTHB)   + "\n\n" +

      "💰 Dividend YTD : " + fmtTHB(d.divTotalTHB);

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendPortfolioSummary", e);
  }
}


 */

/**
 * ====================================================
 * ฟังก์ชั่นนี้ไม่ใช้
 * ====================================================
 * 
// ----------------------------------------
function sendPortfolioDashboard() {
  sendTelegramSafe("⏳ กำลังสร้าง Dashboard...");
  try {
    const d = collectAllData();

    let msg =
      "📋 PORTFOLIO DASHBOARD\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "🌎 มูลค่ารวม : " + fmtTHB(d.totalPortTHB)   + "\n" +
      "💵 Cash รวม  : " + fmtTHB(d.cashTotal)       + "\n" +
      plEmoji(d.totalUnrealTHB) +
        " Unrealized : " + signTHB(d.totalUnrealTHB) + "\n" +
      "💰 Dividend  : " + fmtTHB(d.divTotalTHB)     + "\n" +
      "💱 USD/THB   : " + fmt4(d.fxRate)             + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇹🇭 หุ้นไทย (" + d.thHoldings.length + " ตัว)\n";

    for (const h of d.thHoldings) {
      msg +=
        plEmoji(h.unrealizedPL) + " " + h.ticker +
        " | " + fmtShares(h.sharesRemain) + " หุ้น" +
        " | avg " + fmtTHB(h.avgCost) +
        " | now " + fmtTHB(h.priceNow) +
        " | " + signTHB(h.unrealizedPL) +
        " (" + signPct(h.unrealizedPct) + ")\n";
    }

    msg += "\n━━━━━━━━━━━━\n🇺🇸 หุ้นสหรัฐ (" +
      d.usHoldings.length + " ตัว)\n";

    for (const h of d.usHoldings) {
      msg +=
        plEmoji(h.unrealizedPL) + " " + h.ticker +
        " | " + fmtShares(h.sharesRemain) + " หุ้น" +
        " | avg " + fmtUSD(h.avgCost) +
        " | now " + fmtUSD(h.priceNow) +
        " | " + signUSD(h.unrealizedPL) +
        " (" + signPct(h.unrealizedPct) + ")\n";
    }

    // ✅ เพิ่มกองทุนรวม
    msg += "\n━━━━━━━━━━━━\n🏛️ กองทุนรวม (" +
      d.fundHoldings.length + " กองทุน)\n";

    if (d.fundHoldings.length === 0) {
      msg += "ไม่มีกองทุนที่ถืออยู่\n";
    } else {
      for (const f of d.fundHoldings) {
        msg +=
          plEmoji(f.unrealizedPL) + " " + f.name +
          " | " + fmt(f.unitsRemain) + " หน่วย" +
          " | avg " + fmtTHB(f.avgCost) +
          " | NAV " + fmtTHB(f.navNow) +
          " | " + signTHB(f.unrealizedPL) +
          " (" + signPct(f.unrealizedPct) + ")\n";
      }
    }

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendPortfolioDashboard", e);
  }
}

 */


/**
 * ====================================================
 * ฟังก์ชั่นนี้ไม่ใช้
 * ====================================================
 * 


// ----------------------------------------
function sendPortfolioRanking() {
  sendTelegramSafe("⏳ กำลังรวบรวมข้อมูลหุ้น...");
  try {
    const d = collectAllData();

    const thSorted   = [...d.thHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);
    const usSorted   = [...d.usHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);
    const fundSorted = [...d.fundHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);

    const thTotal   = d.thHoldings.reduce((s, h) => s + h.valueNow, 0);
    const usTotal   = d.usHoldings.reduce((s, h) => s + h.valueNow, 0);
    const fundTotal = d.fundHoldings.reduce((s, h) => s + h.valueNow, 0);

    let msg =
      "🏆 PORTFOLIO RANKING\n" +
      "🕐 " + getNow() + "\n━━━━━━━━━━━━\n\n" +
      "🇹🇭 หุ้นไทย\n";

    thSorted.forEach((h, i) => {
      const w = thTotal > 0 ? (h.valueNow / thTotal) * 100 : 0;
      msg +=
        (i + 1) + ". " + h.ticker + "\n" +
        "  💰 " + fmtTHB(h.avgCost) +
        "  📈 " + fmtTHB(h.priceNow) +
        "  " + plEmoji(h.unrealizedPL) +
        " " + signPct(h.unrealizedPct) +
        "  ⚖️ " + fmt(w) + "%\n";
    });

    msg += "\n━━━━━━━━━━━━\n🇺🇸 หุ้นสหรัฐ\n";

    usSorted.forEach((h, i) => {
      const w = usTotal > 0 ? (h.valueNow / usTotal) * 100 : 0;
      msg +=
        (i + 1) + ". " + h.ticker + "\n" +
        "  💰 " + fmtUSD(h.avgCost) +
        "  📈 " + fmtUSD(h.priceNow) +
        "  " + plEmoji(h.unrealizedPL) +
        " " + signPct(h.unrealizedPct) +
        "  ⚖️ " + fmt(w) + "%\n";
    });

    // ✅ เพิ่มกองทุนรวม
    msg += "\n━━━━━━━━━━━━\n🏛️ กองทุนรวม\n";

    if (fundSorted.length === 0) {
      msg += "ไม่มีกองทุน\n";
    } else {
      fundSorted.forEach((f, i) => {
        const w = fundTotal > 0 ? (f.valueNow / fundTotal) * 100 : 0;
        msg +=
          (i + 1) + ". " + f.name + "\n" +
          "  💰 " + fmtTHB(f.avgCost) +
          "  📈 NAV " + fmtTHB(f.navNow) +
          "  " + plEmoji(f.unrealizedPL) +
          " " + signPct(f.unrealizedPct) +
          "  ⚖️ " + fmt(w) + "%\n";
      });
    }

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendPortfolioRanking", e);
  }
}

 */


// /portfolio = Portfolio Report (เดิมชื่อ sendPortfolioSummary)
function sendPortfolioReport() {
  try {
    sendTelegramSafe("⏳ กำลังสร้าง Portfolio Report...");
    const d = collectAllData();

    let msg =
      "📊 PORTFOLIO REPORT\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "💵 เงินสด (🇹🇭) : " + fmtTHB(d.cashTHB) + "\n" +
      "💵 เงินสด (🇺🇸) : " + fmtUSD(d.cashUSD) +
        " (" + fmtTHB(d.cashUSD * d.fxRate) + ")\n" +
      "💵 เงินสดคงเหลือ : " + fmtTHB(d.cashTotal) + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇺🇸 พอร์ตสหรัฐ\n" +
      "📈 มูลค่า      : " + fmtUSD(d.usPortUSD) +
        " (" + fmtTHB(d.usPortTHB) + ")\n" +
      plEmoji(d.usUnrealUSD) +
        " Unrealized : " + signUSD(d.usUnrealUSD) +
        " (" + signTHB(d.usUnrealTHB) + ")\n" +
      plEmoji(d.usRealUSD) +
        " Realized   : " + signUSD(d.usRealUSD) +
        " (" + signTHB(d.usRealTHB) + ")\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇹🇭 พอร์ตไทย\n" +
      "📈 มูลค่า      : " + fmtTHB(d.thPortTHB) + "\n" +
      plEmoji(d.thUnrealTHB) +
        " Unrealized : " + signTHB(d.thUnrealTHB) + "\n" +
      plEmoji(d.thRealTHB) +
        " Realized   : " + signTHB(d.thRealTHB) + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🏛️ กองทุนรวม\n" +
      "📈 มูลค่า      : " + fmtTHB(d.fundPortTHB) + "\n" +
      plEmoji(d.fundUnrealTHB) +
        " Unrealized : " + signTHB(d.fundUnrealTHB) + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🌎 พอร์ตรวม\n" +
      "📈 มูลค่า      : " + fmtTHB(d.totalPortTHB) + "\n" +
      plEmoji(d.totalUnrealTHB) +
        " Unrealized : " + signTHB(d.totalUnrealTHB) + "\n" +
      plEmoji(d.totalRealTHB) +
        " Realized   : " + signTHB(d.totalRealTHB) + "\n\n" +

      "💰 Dividend YTD : " + fmtTHB(d.divTotalTHB);

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendPortfolioReport", e);
  }
}






// ----------------------------------------
// /weekly
// ----------------------------------------

function sendWeeklyReport() {
  sendTelegramSafe("⏳ กำลังสร้าง Portfolio Report...");
  try {
    const d      = collectAllData();
    const medals = ["🥇", "🥈", "🥉"];

    let msg =
      "📊 PORTFOLIO REPORT\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "💵 เงินสด (🇹🇭)  : " + fmtTHB(d.cashTHB) + "\n" +
      "💵 เงินสด (🇺🇸) : " + fmtUSD(d.cashUSD) +
        " (" + fmtTHB(d.cashUSD * d.fxRate) + ")\n" +
      "💵 เงินสดคงเหลือ : " + fmtTHB(d.cashTotal) + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇺🇸 พอร์ตสหรัฐ\n" +
      "📈 มูลค่า      : " + fmtUSD(d.usPortUSD) +
        " (" + fmtTHB(d.usPortTHB) + ")\n" +
      plEmoji(d.usUnrealUSD) +
        " Unrealized : " + signUSD(d.usUnrealUSD) +
        " (" + signTHB(d.usUnrealTHB) + ")\n" +
      plEmoji(d.usRealUSD) +
        " Realized   : " + signUSD(d.usRealUSD) +
        " (" + signTHB(d.usRealTHB) + ")\n\n" +

      "━━━━━━━━━━━━\n" +
      "🇹🇭 พอร์ตไทย\n" +
      "📈 มูลค่า      : " + fmtTHB(d.thPortTHB)    + "\n" +
      plEmoji(d.thUnrealTHB) +
        " Unrealized : " + signTHB(d.thUnrealTHB)  + "\n" +
      plEmoji(d.thRealTHB) +
        " Realized   : " + signTHB(d.thRealTHB)    + "\n\n" +

      // ✅ เพิ่มกองทุนรวม
      "━━━━━━━━━━━━\n" +
      "🏛️ กองทุนรวม\n" +
      "📈 มูลค่า      : " + fmtTHB(d.fundPortTHB)    + "\n" +
      plEmoji(d.fundUnrealTHB) +
        " Unrealized : " + signTHB(d.fundUnrealTHB)  + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🌎 พอร์ตรวม\n" +
      "📈 มูลค่า      : " + fmtTHB(d.totalPortTHB)    + "\n" +
      plEmoji(d.totalUnrealTHB) +
        " Unrealized : " + signTHB(d.totalUnrealTHB) + "\n" +
      plEmoji(d.totalRealTHB) +
        " Realized   : " + signTHB(d.totalRealTHB)   + "\n\n" +

      "━━━━━━━━━━━━\n🇺🇸 Top US\n";

    d.usTop3.forEach((h, i) => {
      msg += medals[i] + " " + h.ticker +
        "  " + signPct(h.unrealizedPct*100) +
        "  (" + fmtUSD(h.priceNow) + ")\n";
    });
    if (d.usWorst) {
      msg += "📉 แย่สุด " + d.usWorst.ticker +
        "  " + signPct(d.usWorst.unrealizedPct) + "\n";
    }

    msg += "\n━━━━━━━━━━━━\n🇹🇭 Top TH\n";
    d.thTop3.forEach((h, i) => {
      msg += medals[i] + " " + h.ticker +
        "  " + signPct(h.unrealizedPct*100) +
        "  (" + fmtTHB(h.priceNow) + ")\n";
    });
    if (d.thWorst) {
      msg += "📉 แย่สุด " + d.thWorst.ticker +
        "  " + signPct(d.thWorst.unrealizedPct) + "\n";
    }

    // ✅ เพิ่ม กองทุน top
    if (d.fundHoldings.length > 0) {
      msg += "\n━━━━━━━━━━━━\n🏛️ กองทุน\n";
      d.fundHoldings.forEach((f, i) => {
        msg += medals[i] + " " + f.name +
          "  " + signPct(f.unrealizedPct*100) +
          "  (NAV " + fmtTHB(f.navNow) + ")\n";
      });
    }

    msg +=
      "\n━━━━━━━━━━━━\n" +
      "💰 Dividend YTD : " + fmtTHB(d.divTotalTHB) + "\n" +
      "💱 FX Rate      : " + fmt4(d.fxRate);

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendWeeklyReport", e);
  }
}

// ========================================
// Weekly Performance Report
// เปรียบเทียบสัปดาห์นี้ vs สัปดาห์ก่อน
// ========================================
function sendWeeklyPerformance() {
  try {
    sendTelegramSafe("⏳ กำลังสร้าง Weekly Performance...");
    const d     = collectAllData();
    const today = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy");

    const totalReturn    = d.totalUnrealTHB + d.totalRealTHB + d.divTotalTHB;
    const totalReturnPct = d.netInvest > 0
      ? (totalReturn / d.netInvest) * 100 : 0;

    const usSorted = [...d.usHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);
    const thSorted = [...d.thHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);

    let msg =
      "📊 WEEKLY PERFORMANCE\n" +
      "🕐 " + today + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "💹 ผลตอบแทนรวม\n" +
      plEmoji(d.totalUnrealTHB) +
        " Unrealized : " + signTHB(d.totalUnrealTHB) + "\n" +
      plEmoji(d.totalRealTHB) +
        " Realized   : " + signTHB(d.totalRealTHB)   + "\n" +
      "💰 Dividend   : " + fmtTHB(d.divTotalTHB)     + "\n" +
      plEmoji(totalReturn) +
        " Net Return  : " + signTHB(totalReturn)      + "\n" +
      plEmoji(totalReturnPct) +
        " Net Return% : " + signPct(totalReturnPct)   + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "🏆 Top 3 US (P/L%)\n";

    const medals = ["🥇", "🥈", "🥉"];
    usSorted.slice(0, 3).forEach((h, i) => {
      msg += medals[i] + " " + h.ticker +
        " : " + signPct(h.unrealizedPct * 100) +
        " (" + fmtUSD(h.priceNow) + ")\n";
    });
    if (usSorted.length > 0) {
      const w = usSorted[usSorted.length - 1];
      msg += "📉 แย่สุด: " + w.ticker +
        " : " + signPct(w.unrealizedPct * 100) + "\n";
    }

    msg += "\n━━━━━━━━━━━━\n🏆 Top 3 TH (P/L%)\n";
    thSorted.slice(0, 3).forEach((h, i) => {
      msg += medals[i] + " " + h.ticker +
        " : " + signPct(h.unrealizedPct * 100) +
        " (" + fmtTHB(h.priceNow) + ")\n";
    });
    if (thSorted.length > 0) {
      const w = thSorted[thSorted.length - 1];
      msg += "📉 แย่สุด: " + w.ticker +
        " : " + signPct(w.unrealizedPct * 100) + "\n";
    }

    // ✅ กองทุนรวม
    if (d.fundHoldings.length > 0) {
      msg += "\n━━━━━━━━━━━━\n🏛️ กองทุนรวม\n";
      d.fundHoldings.forEach((f, i) => {
        msg += medals[i] + " " + f.name +
          " : " + signPct(f.unrealizedPct * 100) +
          " (NAV ฿" + fmt(f.navNow) + ")\n";
      });
    }

    msg +=
      "\n━━━━━━━━━━━━\n" +
      "💱 USD/THB : " + fmt4(d.fxRate) + "\n" +
      "💰 Dividend YTD : " + fmtTHB(d.divTotalTHB);

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendWeeklyPerformance", e);
  }
}


// ----------------------------------------
// Auto Trigger Weekly
// ----------------------------------------
function sendWeeklyReportAuto() {
  try {
    sendWeeklyReport();
  } catch (e) {
    logError("sendWeeklyReportAuto", e);
  }
}

function createWeeklyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendWeeklyReportAuto")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendWeeklyReportAuto")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .inTimezone("Asia/Bangkok")
    .create();

  Logger.log("✅ Weekly Trigger: ทุกวันจันทร์ 08:00");

}


// ========================================
// /holdings — Holdings by Profit
// ========================================
/**
 * ====================================================
 * ### ข้อควรระวัง
 * เพราะต้องดึง Yahoo ทุกตัว อาจช้าครับ ถ้าหุ้น 13 ตัว × 400ms = ~5 วินาที ควรเพิ่ม loading message ต้น function:

```js
sendTelegramSafe("⏳ กำลังดึงข้อมูล " + 
  (d.thHoldings.length + d.usHoldings.length) + 
  " หุ้น อาจใช้เวลา 10-15 วินาที...");
```
 * ====================================================
 * 
 */
function sendHoldingsByProfit() {
  try {
    const d      = collectAllData();
    const fxRate = d.fxRate;

    // ✅ declare ก่อนใช้ทุกตัว
    const usHoldings   = d.usHoldings;
    const thHoldings   = d.thHoldings;
    const fundHoldings = d.fundHoldings;

    const thSorted = [...thHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct
    );
    const usSorted = [...usHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct
    );

    // ดึง ATR map
    const usATRRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thATRRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);
    const atrMap    = {};
    [...usATRRows, ...thATRRows].forEach(r => {
      atrMap[r.symbol.toUpperCase()] = r;
    });

    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm"
    );

    sendTelegramSafe(
      "⏳ กำลังดึงข้อมูล " +
      (thHoldings.length + usHoldings.length) +
      " หุ้น อาจใช้เวลา 10-15 วินาที..."
    );

    // ── Helper: emoji ──
    const profitEmoji = pct => {
      if (pct >= 10) return "🟢";
      if (pct >= 0)  return "🟡";
      return "🔴";
    };

    // ── Helper: buildCard ──
    const buildCard = (h, idx, cur, isTH) => {
      const pct    = h.unrealizedPct * 100;
      const plSign = h.unrealizedPL >= 0 ? "+" : "";
      const plStr  = plSign + cur + fmt(h.unrealizedPL);
      const emoji  = profitEmoji(pct);
      const atr    = atrMap[h.ticker.toUpperCase()];

      // Yahoo Finance
      let changeP = 0, volRatio = 0, volNow = 0, volAvg = 0;
      try {
        const yahooSymbol = isTH ? h.ticker + ".BK" : h.ticker;
        const stockData   = _fetchStockData(yahooSymbol);
        if (stockData) {
          changeP  = stockData.chgPct   || 0;
          volRatio = stockData.volRatio  || 0;
          volNow   = stockData.volNow    || 0;
          volAvg   = stockData.volAvg20  || 0;
        }
      } catch (e) {
        logError("buildCard [" + h.ticker + "]", e);
      }

      // Price Alert
      const stopLevel  = atr ? atr.trailingStop : 0;
      const distToStop = stopLevel > 0 && h.priceNow > 0
        ? ((h.priceNow - stopLevel) / h.priceNow) * 100 : 0;
      const priceAlert =
        h.priceNow < stopLevel ? "🚨 ต่ำกว่า Stop — ขายทันที!" :
        distToStop < 2         ? "🔴 ใกล้ Stop มาก (<2%)"      :
        distToStop < 5         ? "🟡 ใกล้ Stop (<5%)"          :
                                 "🟢 +" + fmt(distToStop) + "%";

      // Volatility Alert
      const absChange  = Math.abs(changeP);
      const volatAlert =
        absChange >= 3 ? "🔴 ผันผวนสูง " + signPct(changeP)     :
        absChange >= 2 ? "🟡 ผันผวนปานกลาง " + signPct(changeP) :
                         "🟢 ปกติ " + signPct(changeP);

      // Volume Alert
      const volAlert =
        volRatio >= 2.0 ? "🔴 Volume พุ่ง ×" + fmt(volRatio) + " — มีนัยสำคัญ" :
        volRatio >= 1.5 ? "🟡 Volume สูง ×"  + fmt(volRatio)                   :
                          "🟢 ปกติ ×"         + fmt(volRatio);

      // Volume Signal
      let volSignal = "";
      if (changeP < 0 && volRatio >= 1.5) {
        volSignal = "  ⚡ Signal : ⚠️ ราคาลง+Volume สูง — Sell Pressure\n";
      } else if (changeP > 0 && volRatio >= 1.5) {
        volSignal = "  ⚡ Signal : ✅ ราคาขึ้น+Volume สูง — Buy Pressure\n";
      }

      let card =
        idx + "." + emoji + " " + h.ticker +
        "  " + signPct(pct) + " (" + plStr + ")\n" +
        "   🔅 Avg " + cur + fmt(h.avgCost) +
        " → Current " + cur + fmt(h.priceNow) + "\n" +
        "   📦 Qty : " + fmt(h.sharesRemain) + " หุ้น\n" +
        "   💵 มูลค่าปัจจุบัน : " + cur + fmt(h.valueNow) + "\n\n";

      if (atr) {
        card +=
          " ATR Stop : " + cur + fmt(stopLevel) + "\n" +
          "📌 Status  : " + atr.status + "\n\n";
      }

      card +=
        "📊 Alerts\n" +
        "  💰 Price     : " + priceAlert + "\n" +
        "  📈 Volatility: " + volatAlert + "\n" +
        "  📦 Volume    : " + volAlert   + "\n" +
        volSignal;

      return card + "\n";
    };

    // ── Build Message ──
    let msg =
      "📊 HOLDINGS BY PROFIT\n" +
      "🕐 " + today + "\n" +
      "━━━━━━━━━━━━\n\n";

    // TH
    msg += "🇹🇭 หุ้นไทย\n🏆 Top Ranking\n\n";
    thSorted.forEach((h, i) => {
      msg += buildCard(h, i + 1, "฿", true);
    });

    // US
    msg += "━━━━━━━━━━━━\n\n🇺🇸 หุ้นสหรัฐ\n🏆 Top Ranking\n\n";
    usSorted.forEach((h, i) => {
      msg += buildCard(h, i + 1, "$", false);
    });

    // Fund
    if (fundHoldings.length > 0) {
      msg += "━━━━━━━━━━━━\n\n🏛️ กองทุนรวม\n\n";
      fundHoldings.forEach((f, i) => {
        const pct   = f.unrealizedPct * 100;
        const emoji = profitEmoji(pct);
        msg +=
          (i + 1) + "." + emoji + " " + f.name +
          "  " + signPct(pct) + "\n" +
          "   💰 avg cost : ฿" + fmt(f.avgCost) +
          " → NAV : ฿" + fmt(f.navNow) + "\n" +
          "   📦 Units : " + fmt(f.unitsRemain) + " หน่วย\n" +
          "   💵 มูลค่าปัจจุบัน : ฿" + fmt(f.valueNow) + "\n\n";
      });
    }

    // Statistics
    const thProfit = thSorted.filter(h => h.unrealizedPct >= 0).length;
    const thLoss   = thSorted.filter(h => h.unrealizedPct <  0).length;
    const usProfit = usSorted.filter(h => h.unrealizedPct >= 0).length;
    const usLoss   = usSorted.filter(h => h.unrealizedPct <  0).length;
    const fundProfit = fundHoldings.filter(f => f.unrealizedPct >= 0).length;
    const fundLoss   = fundHoldings.filter(f => f.unrealizedPct <  0).length;

    const allSorted = [
      ...thSorted.map(h => ({ ticker: h.ticker, pct: h.unrealizedPct * 100 })),
      ...usSorted.map(h => ({ ticker: h.ticker, pct: h.unrealizedPct * 100 })),
      ...fundHoldings.map(f => ({ ticker: f.name, pct: f.unrealizedPct * 100 }))
    ].sort((a, b) => b.pct - a.pct);

    const best  = allSorted[0];
    const worst = allSorted[allSorted.length - 1];

    msg +=
      "━━━━━━━━━━━━\n\n" +
      "📊 Portfolio Statistics\n\n" +
      "🇹🇭 หุ้นไทย\n" +
      "🟢 กำไร   : " + thProfit + " ตัว\n" +
      "🔴 ขาดทุน : " + thLoss   + " ตัว\n\n" +
      "🇺🇸 หุ้นสหรัฐ\n" +
      "🟢 กำไร   : " + usProfit + " ตัว\n" +
      "🔴 ขาดทุน : " + usLoss   + " ตัว\n\n";

    if (fundHoldings.length > 0) {
      msg +=
        "🏛️ กองทุนรวม\n" +
        "🟢 กำไร   : " + fundProfit + " กองทุน\n" +
        "🔴 ขาดทุน : " + fundLoss   + " กองทุน\n\n";
    }

    msg +=
      "━━━━━━━━━━━━\n" +
      "🏆 Best Performer  : " +
        (best  ? best.ticker  + " (" + signPct(best.pct)  + ")" : "—") + "\n" +
      "📉 Worst Performer : " +
        (worst ? worst.ticker + " (" + signPct(worst.pct) + ")" : "—") + "\n\n" +
      "📌 เงื่อนไขสถานะ\n" +
      "🟢 Strong Profit > 10%\n" +
      "🟡 Small Profit  0-10%\n" +
      "🔴 Loss";

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("sendHoldingsByProfit", e);
  }
}


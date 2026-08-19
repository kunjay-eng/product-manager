// ========================================
// summary_dashboard.gs
// ========================================

function sendSummaryDashboard() {
  try {
    // ── 1. ดึงข้อมูลพื้นฐานก่อน ──
    const fxRate  = getFxRate();
    const cash    = getSheet(SHEETS.CASH);
    const realUS  = getSheet(SHEETS.US_REAL);

    // ── 2. Cash Flow cells ──
    const deposit     = Number(cash.getRange("C6").getValue())  || 0;
    const withdraw    = Number(cash.getRange("C7").getValue())  || 0;
    const exchangeTHB = Number(cash.getRange("C11").getValue()) || 0;
    const cashTHB     = Number(cash.getRange("C14").getValue()) || 0;
    const cashUSD     = Number(cash.getRange("F14").getValue()) || 0;
    const totalUSDIn  = Number(cash.getRange("F6").getValue())  || 0;
    const withdrawUSD = Number(cash.getRange("F7").getValue())  || 0;
    const divTH       = Number(cash.getRange("C10").getValue()) || 0;
    const divUS       = Number(cash.getRange("F10").getValue()) || 0;
    //const feeUS       = Number(cash.getRange("F11").getValue()) || 0;

    // ── 3. Realized P&L ──
    const realTH_THB = Number(realUS.getRange("S2").getValue()) || 0;
    const realUS_USD = Number(realUS.getRange("U2").getValue()) || 0;
    const realTH_tot = Number(realUS.getRange("S3").getValue()) || 0;
    const feeUS = Number(realUS.getRange("O5").getValue()) || 0;



    // ── 4. Holdings — declare ก่อนใช้ ──
    const usHoldings   = getHoldings(SHEETS.US_HOLD);
    const thHoldings   = getHoldings(SHEETS.TH_HOLD);
    const fundHoldings = getFundHoldings();

    // ── 5. คำนวณจาก Holdings ──
    const costTH    = thHoldings.reduce((s, h) => s + h.totalCost, 0);
    const costFund  = fundHoldings.reduce((s, h) => s + h.totalCost, 0);
    const costUS    = usHoldings.reduce((s, h) => s + h.totalCost, 0);

    const thPortTHB   = thHoldings.reduce((s, h) => s + h.valueNow, 0);
    const fundPortTHB = fundHoldings.reduce((s, h) => s + h.valueNow, 0);
    const usPortUSD   = usHoldings.reduce((s, h) => s + h.valueNow, 0);

    const thUnrealTHB   = thHoldings.reduce((s, h) => s + h.unrealizedPL, 0);
    const fundUnrealTHB = fundHoldings.reduce((s, h) => s + h.unrealizedPL, 0);
    const usUnrealUSD   = usHoldings.reduce((s, h) => s + h.unrealizedPL, 0);

    // ── 6. มูลค่าสินทรัพย์รวม ──
    //const totalAssetTHB = thPortTHB + cashTHB + fundPortTHB;
    //const totalAssetUSD = usPortUSD + cashUSD;
    const totalAssetTHB = thPortTHB ;
    const totalAssetUSD = usPortUSD ;
    const totalAssetCashUSD = cashUSD ;
    const totalAssetCashTHB = cashTHB ;


    
    // ── 7. เงินลงทุนสุทธิ ──
    const netInvestTHB = deposit - withdraw - exchangeTHB;
    const netInvestUSD = totalUSDIn - withdrawUSD;

    // ── 8. Dividend YTD ──
    const divYTD  = getDividendYTD();
    const divTotal = divYTD.totalTHB;

    // ── 9. ค่าธรรมเนียม ──
    const feeTH    = _getFeeTH();
    const feeUSThb = feeUS * fxRate;
    const feeTotal = feeTH + feeUSThb;

    // ── 10. ผลตอบแทนรวม ──
    const totalUnreal  = thUnrealTHB + (usUnrealUSD * fxRate) + fundUnrealTHB;
    const totalReturn  = totalUnreal + realTH_tot + divTotal;
    const netInvestAll = netInvestTHB + (netInvestUSD * fxRate);
    const returnPct    = netInvestAll > 0
      ? (totalReturn / netInvestAll) * 100 : 0;
    const netReturn    = totalReturn - feeTotal;
    const netReturnPct = netInvestAll > 0
      ? (netReturn / netInvestAll) * 100 : 0;

    // ── 11. สร้าง message ──
    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss"
    );

    let msg =
      "\n" +

      "📊 Dime! Summary\n" +
      today + "\n\n" +

      "💵 เงินฝาก           : " + fmtTHB(deposit)     + "\n" +
      "💸 เงินถอน           : " + fmtTHB(withdraw)    + "\n" +
      "💱 แลกเงิน THB→USD  : " + fmtTHB(exchangeTHB) + "\n" +
      "─────────────────────\n\n" +

      "🇹🇭  หุ้นไทย\n" +
      "💰 ลงทุนสุทธิ     : " + fmtTHB(netInvestTHB)       + "\n" +
      "📦 ต้นทุนรวม     : " + fmtTHB(costTH + costFund)  + "\n" +
      "💰 Cash        : " + fmtTHB(cashTHB)            + "\n" +
      "📈 มูลค่าสินทรัพย์รวม: " + fmtTHB(totalAssetTHB)      + "\n\n" +

      "🔵 กำไรทิพย์ (Unrealized)\n" +
      "   ยังไม่ขาย ราคาตลาดเปลี่ยนได้\n" +
      "   " + plEmoji(thUnrealTHB) + " " + signTHB(thUnrealTHB) + "\n\n" +

      "🟡 กำไรจริง (Realized)\n" +
      "   ขายแล้ว ได้รับจริง\n" +
      "   " + plEmoji(realTH_THB) + " " + signTHB(realTH_THB)   + "\n\n" +

      "💰 ปันผลสะสม   : " + signTHB(divTH)  + "\n" +
      "💸 Fee    : " + fmtTHB(feeTH)   + "\n" +
      "─────────────────────\n\n" +

      "🇺🇸 หุ้นสหรัฐ\n" +
      "💰 ลงทุนสุทธิ    : " + fmtUSD(netInvestUSD)  + "\n" +
      "📦 ต้นทุนรวม    : " + fmtUSD(costUS)         + "\n" +
      "💰 Cash      : " + signUSD(cashUSD)        + "\n" +
      "📈 มูลค่าสินทรัพย์รวม: " + fmtUSD(totalAssetUSD)  + "\n" +
      "   คิดเป็นเงินไทย  ("+ signTHB(totalAssetUSD * fxRate)  + ") \n\n" +

      "🔵 กำไรทิพย์ (Unrealized)\n" +
      "   ยังไม่ขาย ราคาตลาดเปลี่ยนได้\n" +
      "   " + plEmoji(usUnrealUSD) + " " + signUSD(usUnrealUSD)  + " ("+ signTHB(usUnrealUSD * fxRate)  + ")\n\n" +

      "🟡 กำไรจริง (Realized)\n" +
      "   ขายแล้ว ได้รับจริง\n" +
      "   " + plEmoji(realUS_USD) + " " + signUSD(realUS_USD)  + " ("+ signTHB(realUS_USD * fxRate)  + ")\n\n" +

      "💰 ปันผลสะสม         : " + signUSD(divUS) + "\n" +
      "💸 Fee            : " + fmtUSD(feeUS)  + "\n" +
      "─────────────────────\n\n";

    // กองทุนรวม
    if (fundHoldings.length > 0) {
      msg += "🏛️ กองทุนรวม\n";
      fundHoldings.forEach(f => {
        msg +=
          "  " + f.name + "\n" +
          "  NAV " + fmtTHB(f.navNow) +
          "  |  P&L " + signTHB(f.unrealizedPL) +
          " (" + signPct(f.unrealizedPct * 100) + ")\n";
      });
      msg +=
        "📈 มูลค่าสินทรัพย์รวม: " + fmtTHB(fundPortTHB) + "\n" +
        "─────────────────────\n";
    }

      msg +=
      "\n" +
      "📊 สรุปผลตอบแทนรวม\n\n" +

      "📈 มูลค่าสินทรัพย์รวม: " + fmtTHB(fundPortTHB+(totalAssetUSD * fxRate)+totalAssetTHB+(cashTHB+cashUSD*fxRate)) + "\n\n" +
      "🔵 Unrealized \n" +
      "ยังไม่ขาย ราคาตลาดเปลี่ยนได้\n" +
      "  🇹🇭 หุ้นไทย    : " + signTHB(thUnrealTHB)          + "\n" +
      "  🇺🇸 หุ้นสหรัฐ    : " + signTHB(usUnrealUSD * fxRate)  + "\n" +
      "  🏛️ กองทุน   : " + signTHB(fundUnrealTHB)         + "\n\n" +

      "🔵 กำไรทิพย์ (Unrealized) :" + plEmoji(totalUnreal) + " " + signTHB(totalUnreal)  + "\n\n" +

     // "  รวมกำไรทิพย์  : " + signTHB(totalUnreal)            + "\n\n" +

      "🟡 Realized  \n" +
      "ขายแล้ว ได้รับจริง\n" +
      "  🇹🇭 หุ้นไทย   : " + signTHB(realTH_THB)            + "\n" +
      "  🇺🇸 หุ้นสหรัฐ    : " + signTHB(realUS_USD * fxRate)   + "\n\n" +
      "🟡 กำไรจริง (Realized) : " + plEmoji(realTH_tot) + " " + signTHB(realTH_tot)  + "\n\n" + 
      
     // "  รวมกำไรจริง   : " + signTHB(realTH_tot)             + "\n\n" +

      "  💰 ปันผล YTD     : " + signTHB(divTotal)       + "\n" +
      "  💸 ค่าธรรมเนียม       : " + fmtTHB(feeTotal)        + "\n\n" +
      "อัตราแลกเปลี่ยน : " + fmtTHB(fxRate)  + "\n" +
      "─────────────────────\n" +

      plEmoji(totalReturn) +
        " ผลตอบแทนรวม  : " + signTHB(totalReturn) +
        " (" + signPct(returnPct) + ")\n" ;
     // plEmoji(netReturn) ;
     //   " ผลตอบแทนรวม หัก fee : " + signTHB(netReturn) +
    //    " (" + signPct(netReturnPct) + ")";


    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("sendSummaryDashboard", e);
  }
}

// ----------------------------------------
// Auto Trigger — ทุกวัน 19:00
// ----------------------------------------
function sendSummaryDashboardAuto() {
  try {
    sendSummaryDashboard();
  } catch (e) {
    logError("sendSummaryDashboardAuto", e);
  }
}

// ----------------------------------------
// Setup Trigger
// ----------------------------------------
function createSummaryDashboardTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t =>
      t.getHandlerFunction() === "sendSummaryDashboardAuto"
    )
    .forEach(t => ScriptApp.deleteTrigger(t));


  ScriptApp.newTrigger("sendSummaryDashboardAuto")
    .timeBased()
    .everyDays(1)
    .atHour(19)
    .inTimezone("Asia/Bangkok")
    .create();


  Logger.log("✅ Summary Dashboard Trigger: ทุกวัน 19:00");
}


// ── ดึงค่าธรรมเนียม TH ──
function _getFeeTH() {
  try {
    const sheet   = getSheet(SHEETS.TH_REAL);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.REALIZED) return 0;

    const rows = sheet.getRange(
      START_ROW.REALIZED, 10,
      lastRow - START_ROW.REALIZED + 1, 1
    ).getValues();

    return rows.reduce((s, r) => s + (Number(r[0]) || 0), 0);
  } catch (e) {
    logError("_getFeeTH", e);
    return 0;
  }
}


// ดึงค่าธรรมเนียมจาก Realized P&L
function _getFee() {
  try {
    const usSheet = getSheet(SHEETS.US_REAL);
    const thSheet = getSheet(SHEETS.TH_REAL);
    const lastRowUS = usSheet.getLastRow();
    const lastRowTH = thSheet.getLastRow();

    let feeUSD = 0, feeTHB = 0;

    // US: col J = ค่าธรรมเนียมรวม (USD)
    if (lastRowUS >= START_ROW.REALIZED) {
      const usRows = usSheet.getRange(
        START_ROW.REALIZED, 10,
        lastRowUS - START_ROW.REALIZED + 1, 1
      ).getValues();
      usRows.forEach(r => { feeUSD += Number(r[0]) || 0; });
    }

    // TH: col J = ค่าธรรมเนียมรวม (THB)
    if (lastRowTH >= START_ROW.REALIZED) {
      const thRows = thSheet.getRange(
        START_ROW.REALIZED, 10,
        lastRowTH - START_ROW.REALIZED + 1, 1
      ).getValues();
      thRows.forEach(r => { feeTHB += Number(r[0]) || 0; });
    }

    const fxRate = getFxRate();
    return feeTHB + (feeUSD * fxRate);
  } catch (e) {
    logError("_getFee", e);
    return 0;
  }
}


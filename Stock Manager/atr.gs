// ========================================
// atr.gs
// ATR Trailing Stop
// /atr — ดูทุกตัว
// /atralert — ดูเฉพาะ triggered
// ========================================

// ----------------------------------------
// Data Layer
// ----------------------------------------
function getATRRows(sheetName, startRow) {
  const sheet   = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const numRows = lastRow - startRow + 1;
  const rows    = sheet.getRange(startRow, 1, numRows, 12).getValues();
  const result  = [];

  rows.forEach((row, i) => {
    const symbol = row[ATR_COL.SYMBOL];
    if (!symbol) return;

    result.push({
      rowNumber:    startRow + i,
      symbol:       String(symbol),
      tradeStyle:   String(row[ATR_COL.TRADE_STYLE])  || "",
      buyPrice:     Number(row[ATR_COL.BUY_PRICE])     || 0,
      todayHigh:    Number(row[ATR_COL.TODAY_HIGH])    || 0,
      atr:          Number(row[ATR_COL.ATR])            || 0,
      multiplier:   Number(row[ATR_COL.MULTIPLIER])    || 0,
      initialCL:    Number(row[ATR_COL.INITIAL_CL])    || 0,
      trailingStop: Number(row[ATR_COL.TRAILING_STOP]) || 0,
      status:       String(row[ATR_COL.STATUS])         || "",
      minProfit:    Number(row[ATR_COL.MIN_PROFIT])     || 0,
      shares:       Number(row[ATR_COL.SHARES])         || 0,
      ePricenowATR : Number(row[ATR_COL.Pricenow_ATR])         || 0
    });
  });

  return result;
}

// ----------------------------------------
// Status Helper
// ----------------------------------------
function _atrEmoji(status) {
  if (status.includes("CUT") || status.includes("TAKE PROFIT")) return "🚨";
  if (status.includes("กำลังรันเทรนด์")) return "🔄";
  if (status.includes("กำไร"))          return "🔒";
  return "⚪";
}

function _atrIsAlert(status) {
  return status.includes("CUT") || status.includes("TAKE PROFIT");
}

// ----------------------------------------
// /atr — ดูทุกตัว
// ----------------------------------------
function sendATRAll() {
  sendTelegramSafe("⏳ กำลังดึงข้อมูล ATR Trailing Stop...");
  try {
    _sendATRReport(false);
  } catch (e) {
    sendTelegramError("sendATRAll", e);
  }
}

// ----------------------------------------
// /atralert — ดูเฉพาะตัวที่ triggered
// ----------------------------------------
function sendATRAlert() {
  try {
    _sendATRReport(true);
  } catch (e) {
    sendTelegramError("sendATRAlert", e);
  }
}

// ----------------------------------------
// Auto Trigger — ทุก 30 นาที
// ส่งเฉพาะตัวที่ triggered
// ----------------------------------------
function checkATRAuto() {
  try {
    const usRows = getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);
    const all    = [...usRows, ...thRows];
    const alert  = all.filter(r => _atrIsAlert(r.status));


    if (alert.length === 0) return;

    let msg =
      "🚨 ATR TRAILING STOP ALERT!\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";

    for (const r of alert) {
      const currency = usRows.includes(r) ? "$" : "฿";
      const plPct    = r.buyPrice > 0
        ? ((r.todayHigh - r.buyPrice) / r.buyPrice) * 100 : 0;

      msg +=
        "🚨 " + r.symbol + "\n" +
        "  💰 ต้นทุน avg     : " + currency + fmt(r.buyPrice)     + "\n" +
        "  📈 *ราคาปัจจุบัน*  : " + currency + fmt(r.ePricenowATR)                    + "\n" +
        "  📈 *Today's High*  : " + currency + fmt(r.todayHigh)    + "\n" +
        "  🛑 *Trailing Stop* : " + currency + fmt(r.trailingStop) + "\n"+
        "  💵 กำไรหากขายที่ระดับ Stop "   + "\n" +
        "  📌 สถานะ    : " + r.status     + "\n\n" ;
    }

    sendTelegramSafe(msg);
  } catch (e) {
    logError("checkATRAuto", e);
  }
}

// ----------------------------------------
// Private — Build Full ATR Report
// ----------------------------------------
// เพิ่ม priceNow จาก Holdings

function _sendATRReport(alertOnly) {
  const fxRate = getFxRate();  // ✅ ดึง FX Rate ก่อน

  const sheets = [
    { name: ATR_SHEETS.US, startRow: ATR_START_ROW.US,
      label: "🇺🇸 หุ้นสหรัฐ", currency: "$", isTH: false },
    { name: ATR_SHEETS.TH, startRow: ATR_START_ROW.TH,
      label: "🇹🇭 หุ้นไทย",   currency: "฿", isTH: true  }
  ];

  const thPriceMap = {};
  const usPriceMap = {};
  getHoldings(SHEETS.TH_HOLD).forEach(h => { thPriceMap[h.ticker] = h.priceNow; });
  getHoldings(SHEETS.US_HOLD).forEach(h => { usPriceMap[h.ticker] = h.priceNow; });

  let hasAny = false;

  sheets.forEach(({ name, startRow, label, currency, isTH }) => {
    let rows = getATRRows(name, startRow);
    if (alertOnly) rows = rows.filter(r => _atrIsAlert(r.status));
    if (rows.length === 0) return;

    hasAny = true;

    let msg =
      (alertOnly ? "🚨 ATR ALERT\n" : "📊 ATR TRAILING STOP\n") +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n" +
      label + " (" + rows.length + " ตัว)\n\n";

    for (const r of rows) {
      const priceNow = isTH
        ? (thPriceMap[r.symbol] || 0)
        : (usPriceMap[r.symbol] || 0);

      const plPct = r.buyPrice > 0 && priceNow > 0
        ? ((priceNow - r.buyPrice) / r.buyPrice) * 100 : 0;

      const distStop = r.trailingStop > 0 && r.todayHigh > 0
        ? ((r.todayHigh - r.trailingStop) / r.trailingStop) * 100 : 0;

      // ✅ แสดง (฿xxx) เฉพาะ US เท่านั้น TH ไม่ต้องแปลง
      let minProfitStr;
      if (isTH) {
        minProfitStr = r.minProfit >= 0
          ? currency + fmt(r.minProfit)
          : "-" + currency + fmt(Math.abs(r.minProfit));
      } else {
        const minProfitTHB = r.minProfit * fxRate;
        minProfitStr = r.minProfit >= 0
          ? "$" + fmt(r.minProfit) + " (฿" + fmt(minProfitTHB) + ")"
          : "-$" + fmt(Math.abs(r.minProfit)) + " (-฿" + fmt(Math.abs(minProfitTHB)) + ")";
      }

      const priceNowStr = priceNow > 0 ? currency + fmt(priceNow) : "-";

      msg +=
        _atrEmoji(r.status) + " " + r.symbol + "\n" +
        "  💰 ต้นทุน avg   : " + currency + fmt(r.buyPrice)     + "\n" +
        "  📈 *ราคาปัจจุบัน*  : " + priceNowStr                    + "\n" +
        "  📊 P/L   : " + signPct(plPct) + "\n" + 
        "\n" +         
        "  📈 *Today's High*   : " + currency + fmt(r.todayHigh)    + "\n" +
        "  📉 Initial CL     : " + currency + fmt(r.initialCL)    + "\n" +
        "  🔒 *Trailing Stop*  : " + currency + fmt(r.trailingStop) + "\n" +
        "  📏 ห่าง Trailing  : " + signPct(distStop)              + "\n" +
        "  💵 กำไรหากขายที่ระดับ Stop  : " + minProfitStr                   + "\n" +
        "  ⚙️ ATR (14)     : " + fmt(r.atr) + "  ×" + r.multiplier + "\n" +
        "  📌 สถานะ    : " + r.status                       + "\n"+
        "━━━━━━━━━━━━\n\n"  ; 
    }


    sendTelegramSafe(msg);
  });

  if (!hasAny && alertOnly) {
    sendTelegramSafe(
      "✅ ATR ALERT CHECK\n\n" +
      "ไม่มีหุ้นที่ถูก Trigger\n" +
      "🕐 " + getNow()
    );
  }
}



// ----------------------------------------
// ATR Trigger
// ----------------------------------------
// 

function createATRTrigger() {
  ["checkATRTriggerAuto", "updateATRUSHigh"].forEach(fnName => {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === fnName)
      .forEach(t => ScriptApp.deleteTrigger(t));
  });

  // ✅ ใช้ checkATRTriggerAuto แทน checkATRAuto
  ScriptApp.newTrigger("checkATRTriggerAuto")
    .timeBased().everyMinutes(30).create();

  ScriptApp.newTrigger("updateATRUSHigh")
    .timeBased().everyMinutes(15).create();

  Logger.log("✅ ATR Trigger: check 30 นาที, US High 15 นาที");
}







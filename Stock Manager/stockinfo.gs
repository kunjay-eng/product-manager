// ========================================
// stockinfo.gs
// พิมพ์ชื่อหุ้น → วิเคราะห์ Technical
// ========================================


// ----------------------------------------
// Main — รับ ticker แล้วส่ง analysis
// ----------------------------------------
function sendStockInfo(ticker, manualTarget, manualBudget) {
  try {
    sendTelegramSafe("⏳ กำลังดึงข้อมูล " + ticker + "...");

    const data = _fetchStockData(ticker);
    if (!data) {
      sendTelegramSafe(
        "❌ ไม่พบข้อมูล: " + ticker + "\n" +
        "กรุณาตรวจสอบชื่อหุ้นอีกครั้ง"
      );
      return;
    }

    let msg = _buildStockInfoMsg(ticker, data);

    const targets = _getDCATargets();
    const key     = ticker.replace(".BK", "").toUpperCase();

    if (manualTarget > 0) {
      // ✅ มีการส่ง target/budget มาโดยตรง — ใช้เลย ไม่ต้องดู Sheet
      msg += _appendDCAAnalysisManual(ticker, data, manualTarget, manualBudget || 0);
    } else if (targets[key]) {
      // มีในพอร์ต → DCA Analysis แบบเต็ม
      msg += _appendDCAAnalysis(ticker, data);
    } else {
      // ไม่มีในพอร์ต → Entry Analysis แบบง่าย
      msg += _appendEntryAnalysis(ticker, data);
    }

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("sendStockInfo [" + ticker + "]", e);
  }
}



// ----------------------------------------
// ดึงข้อมูลจาก Yahoo Finance
// ----------------------------------------
function _fetchStockData(ticker) {
  try {
    // ดึง 200 วัน เพื่อคำนวณ MA200
    const url = "https://query1.finance.yahoo.com/v8/finance/chart/" +
      ticker + "?interval=1d&range=1y";

//เช็ค cache ก่อนยิงเอง
const cacheKey = 'yh_' + ticker + '_1y';
let contentText = CacheService.getScriptCache().get(cacheKey);
if (!contentText) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0" } });
  if (resp.getResponseCode() !== 200) return null;
  contentText = resp.getContentText();
} 
  const json = JSON.parse(contentText);

    const result = json?.chart?.result?.[0];
    if (!result) return null;


    const meta   = result.meta;
    const closes = result.indicators?.quote?.[0]?.close  || [];
    const volumes = result.indicators?.quote?.[0]?.volume || [];
    const highs  = result.indicators?.quote?.[0]?.high   || [];
    const lows   = result.indicators?.quote?.[0]?.low    || [];


    // กรองค่า null ออก
    const validCloses = closes.filter(v => v !== null && v !== undefined);
    if (validCloses.length < 20) return null;


    const price  = meta.regularMarketPrice;
    const prev   = meta.previousClose || validCloses[validCloses.length - 2];
    const change = price - prev;
    const chgPct = prev > 0 ? (change / prev) * 100 : 0;


    // คำนวณ MA
    const ma20  = _calcMA(validCloses, 20);
    const ma50  = _calcMA(validCloses, 50);
    const ma200 = _calcMA(validCloses, 200);


    // RSI 14
    const rsi = _calcRSI(validCloses, 14);


    // Volume
    const validVols  = volumes.filter(v => v !== null && v !== undefined);
    const volNow     = validVols[validVols.length - 1] || 0;
    const volAvg20   = _calcMAArr(validVols, 20);
    const volRatio   = volAvg20 > 0 ? volNow / volAvg20 : 0;


    // 52W High/Low
    const w52High = meta.fiftyTwoWeekHigh || Math.max(...validCloses);
    const w52Low  = meta.fiftyTwoWeekLow  || Math.min(...validCloses);


    // ชื่อบริษัท
    const companyName = meta.shortName || meta.longName || ticker;
    const currency    = meta.currency  || "USD";
    const isTH        = ticker.endsWith(".BK");


    return {
      ticker, companyName, currency, isTH,
      price, prev, change, chgPct,
      ma20, ma50, ma200,
      rsi, volNow, volAvg20, volRatio,
      w52High, w52Low
    };


  } catch (e) {
    logError("_fetchStockData [" + ticker + "]", e);
    return null;
  }
}


// ----------------------------------------
// วิเคราะห์ Trend
// ----------------------------------------
function _analyzeTrend(price, ma20, ma50, ma200) {
  if (price > ma20 && ma20 > ma50 && ma50 > ma200) {
    return { label: "🟢 Strong Uptrend", short: "BULL" };
  } else if (price > ma50 && ma50 > ma200) {
    return { label: "🟡 Uptrend", short: "BULL" };
  } else if (price < ma20 && ma20 < ma50 && ma50 < ma200) {
    return { label: "🔴 Strong Downtrend", short: "BEAR" };
  } else if (price < ma50 && ma50 < ma200) {
    return { label: "🔴 Downtrend", short: "BEAR" };
  } else if (price > ma200) {
    return { label: "🟡 Above MA200 (Mixed)", short: "MIXED" };
  } else {
    return { label: "⚪ Sideway / Unclear", short: "SIDE" };
  }
}


// ----------------------------------------
// Signals
// ----------------------------------------
function _buildSignals(price, ma20, ma50, ma200, rsi, volRatio) {
  const signals = [];
  const actions = [];


  // MA signals
  if (price > ma20)  signals.push("✅ Price > MA20");
  else               signals.push("❌ Price < MA20");


  if (ma20 > ma50)   signals.push("✅ MA20 > MA50");
  else               signals.push("❌ MA20 < MA50");


  if (ma50 > ma200)  signals.push("✅ MA50 > MA200");
  else               signals.push("❌ MA50 < MA200");


  // RSI signal
  if (rsi > 70)      signals.push("⚠️ RSI " + fmt(rsi) + " (Overbought)");
  else if (rsi < 30) signals.push("💡 RSI " + fmt(rsi) + " (Oversold)");
  else               signals.push("✅ RSI " + fmt(rsi) + " (Normal)");


  // Volume signal
  if (volRatio > 2)      signals.push("🔥 Volume สูงกว่าปกติ ×" + fmt(volRatio));
  else if (volRatio > 1.5) signals.push("📈 Volume สูงขึ้น ×" + fmt(volRatio));
  else                   signals.push("📊 Volume ปกติ ×" + fmt(volRatio));


  // Action
  if (price > ma20 && ma20 > ma50 && ma50 > ma200) {
    if (rsi > 70) actions.push("⚠️ Overbought — Wait for Pullback");
    else          actions.push("✅ Hold / Add on Dip");
  } else if (price < ma20 && rsi < 40) {
    actions.push("👀 Watch — Potential Reversal");
  } else if (price < ma50 && ma50 < ma200) {
    if (rsi < 30) actions.push("⚠️ Oversold — Risky Buy");
    else          actions.push("🚫 Avoid / Cut Loss");
  } else {
    actions.push("⏳ Wait for Clear Signal");
  }


  return { signals, actions };
}


// ----------------------------------------
// Build Message
// ----------------------------------------
function _buildStockInfoMsg(ticker, d) {
  const cur     = d.isTH ? "฿" : "$";
  const trend   = _analyzeTrend(d.price, d.ma20, d.ma50, d.ma200);
  const { signals, actions } = _buildSignals(
    d.price, d.ma20, d.ma50, d.ma200, d.rsi, d.volRatio
  );


  // P/L เทียบกับ 52W
  const fromHigh = d.w52High > 0
    ? ((d.price - d.w52High) / d.w52High) * 100 : 0;
  const fromLow  = d.w52Low  > 0
    ? ((d.price - d.w52Low)  / d.w52Low)  * 100 : 0;


  const chgEmoji = d.change >= 0 ? "📈" : "📉";
  const chgSign  = d.change >= 0 ? "+" : "";


  let msg =
    chgEmoji + " " + ticker + "\n" +
    "(" + d.companyName + ")\n" +
    "━━━━━━━━━━━━\n\n" +


    "💰 Price      : " + cur + fmt(d.price) +
      "  (" + chgSign + fmt(d.chgPct) + "%)\n\n" +


    "📊 Moving Average\n" +
    "  MA20  : " + cur + fmt(d.ma20)  + "\n" +
    "  MA50  : " + cur + fmt(d.ma50)  + "\n" +
    "  MA200 : " + cur + fmt(d.ma200) + "\n\n" +


    "📉 RSI(14) : " + fmt(d.rsi) + "\n" +
    "📦 Volume  : " + _fmtVol(d.volNow) +
      " (avg20: " + _fmtVol(d.volAvg20) + ")\n\n" +


    "📅 52W Range\n" +
    "  High : " + cur + fmt(d.w52High) +
      " (" + fmt(fromHigh) + "%)\n" +
    "  Low  : " + cur + fmt(d.w52Low)  +
      " (+" + fmt(fromLow) + "%)\n\n" +


    "━━━━━━━━━━━━\n" +
    "🎯 Trend : " + trend.label + "\n\n" +


    "📋 Signal\n";


  signals.forEach(s => { msg += s + "\n"; });


  msg += "\n⚡ Action\n";
  actions.forEach(a => { msg += a + "\n"; });


  return msg;
}


// ----------------------------------------
// คำนวณ MA
// ----------------------------------------
function _calcMA(arr, period) {
  const valid = arr.filter(v => v !== null && v !== undefined);
  if (valid.length < period) return 0;
  const slice = valid.slice(-period);
  return slice.reduce((s, v) => s + v, 0) / period;
}


function _calcMAArr(arr, period) {
  const valid = arr.filter(v => v !== null && v !== undefined);
  if (valid.length < period) return 0;
  const slice = valid.slice(-period);
  return slice.reduce((s, v) => s + v, 0) / period;
}



/**
 * ====================================================
 * Error
 * ====================================================
  */
// ----------------------------------------
// Format Volume
// ----------------------------------------
function _fmtVol(vol) {
  if (vol >= 1000000) return fmt(vol / 1000000) + "M";
  if (vol >= 1000)     return fmt(vol / 1000)     + "K";
  return String(Math.round(vol));
}


// ----------------------------------------
// ซื้อโดยไม่รู้จำนวนหุ้น — ใช้ budget คำนวณให้
// รูปแบบ:
//   VOO 650          → ดึง budget จาก Holdings col Q
//   VOO 650 5000     → ใช้ budget $5,000 ที่ระบุ
// ----------------------------------------
function sendBuyByBudget(text) {
  try {
    const parts      = text.trim().split(/\s+/);
    const ticker     = parts[0].toUpperCase();
    const buyPrice   = parseFloat(parts[1]);
    const inputBudget = parts[2] ? parseFloat(parts[2]) : 0;

    if (isNaN(buyPrice) || buyPrice <= 0) {
      sendTelegramSafe(
        "❌ ราคาซื้อไม่ถูกต้อง: " + parts[1] + "\n\n" +
        "📌 ตัวอย่าง:\n" +
        "VOO 650          (ดึงงบจาก Sheet)\n" +
        "VOO 650 5000     (งบ $5,000)"
      );
      return;
    }

    sendTelegramSafe("⏳ กำลังคำนวณ " + ticker + "...");

    const key  = ticker.replace(".BK", "").toUpperCase();
    const isTH = _isTHStock(key);
    const cur  = isTH ? "฿" : "$";
    const hold = _getCurrentHolding(key, isTH);
    const data = _fetchStockData(isTH ? key + ".BK" : key);

    if (!data) {
      sendTelegramSafe("❌ ไม่พบข้อมูลราคา: " + ticker);
      return;
    }

    // ดึง budget
    // 1. ถ้าระบุมาใน command → ใช้เลย
    // 2. ถ้าไม่ระบุ → ดึงจาก Holdings col Q
    let budget     = inputBudget;
    let budgetSrc  = "ระบุเอง";

    if (budget <= 0) {
      budget    = _getDCABudget(key, isTH);
      budgetSrc = "Holdings col Q";
    }

    if (budget <= 0) {
      sendTelegramSafe(
        "❌ ไม่พบงบประมาณสำหรับ " + ticker + "\n\n" +
        "💡 วิธีแก้:\n" +
        "1. ระบุงบตรงๆ: " + ticker + " " + buyPrice + " 5000\n" +
        "2. กรอกงบใน Holdings col Q"
      );
      return;
    }

    // คำนวณจำนวนหุ้น
    const buyShares   = budget / buyPrice;

    // ข้อมูลเดิม
    const oldShares   = hold ? hold.sharesRemain : 0;
    const oldAvg      = hold ? hold.avgCost : 0;
    const oldCost     = oldShares * oldAvg;

    // หลังซื้อเพิ่ม
    const totalShares  = oldShares + buyShares;
    const totalCost    = oldCost + budget;
    const newAvg       = totalShares > 0 ? totalCost / totalShares : buyPrice;

    const currentPrice = data.price;
    const newValueNow  = totalShares * currentPrice;
    const newUnrealPL  = newValueNow - totalCost;
    const newUnrealPct = totalCost > 0 ? (newUnrealPL / totalCost) * 100 : 0;
    const distBreakEven = newAvg > 0
      ? ((currentPrice - newAvg) / newAvg) * 100 : 0;

    let msg =
      "📊 คาดการณ์ avg cost — " + ticker + "\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";

    // ก่อนซื้อเพิ่ม
    if (hold && oldShares > 0) {
      const oldPLPct = oldAvg > 0
        ? ((currentPrice - oldAvg) / oldAvg) * 100 : 0;
      const oldPL    = oldShares * (currentPrice - oldAvg);

      msg +=
        "📦 ก่อนซื้อเพิ่ม\n" +
        "  ถือ        : " + fmt(oldShares) + " หุ้น\n" +
        "  Avg Cost   : " + cur + fmt(oldAvg) + "\n" +
        "  มูลค่าเดิม : " + cur + fmt(oldShares * currentPrice) + "\n" +
        "  P/L เดิม   : " + signPct(oldPLPct) +
          " (" + (oldPL >= 0 ? "+" : "") + cur + fmt(Math.abs(oldPL)) + ")\n\n";
    } else {
      msg += "📦 ยังไม่มีหุ้นตัวนี้ในพอร์ต\n\n";
    }

    // คำสั่งซื้อ
    msg +=
      "🛒 คำสั่งซื้อ (คำนวณจากงบ)\n" +
      "  ราคาซื้อ : " + cur + fmt(buyPrice) + "\n" +
      "  งบที่ใช้  : " + cur + fmt(budget) +
        " (จาก " + budgetSrc + ")\n" +
      "  ซื้อได้   : " + fmt(buyShares) + " หุ้น\n\n";

    // หลังซื้อเพิ่ม
    msg +=
      "✅ หลังซื้อเพิ่ม\n" +
      "  ถือทั้งหมด : " + fmt(totalShares) + " หุ้น\n" +
      "  ต้นทุนรวม  : " + cur + fmt(totalCost) + "\n";

    if (hold && oldAvg > 0) {
      const avgDiff    = newAvg - oldAvg;
      const avgDiffPct = oldAvg > 0 ? (avgDiff / oldAvg) * 100 : 0;
      const avgEmoji   = avgDiff <= 0 ? "✅" : "⚠️";
      const avgLabel   = avgDiff <= 0 ? "ลดต้นทุน" : "เพิ่มต้นทุน";

      msg +=
        "  Avg Cost เดิม : " + cur + fmt(oldAvg) + "\n" +
        "  Avg Cost ใหม่ : " + cur + fmt(newAvg) + "\n" +
        "  " + avgEmoji + " " + avgLabel + "  " +
          (avgDiff <= 0 ? "-" : "+") + cur + fmt(Math.abs(avgDiff)) +
          " (" + (avgDiff <= 0 ? "" : "+") + fmt(avgDiffPct) + "%)\n\n";
    } else {
      msg +=
        "  Avg Cost ใหม่ : " + cur + fmt(newAvg) + "\n\n";
    }

    // P/L ณ ราคาตลาด
    msg +=
      "📈 P/L ณ ราคาตลาด " + cur + fmt(currentPrice) + "\n" +
      "  มูลค่าพอร์ต  : " + cur + fmt(newValueNow)   + "\n" +
      plEmoji(newUnrealPL) +
        " Unrealized  : " + (newUnrealPL >= 0 ? "+" : "") +
        cur + fmt(Math.abs(newUnrealPL)) +
        " (" + signPct(newUnrealPct) + ")\n" +
      "  Break-even  : " + signPct(distBreakEven) +
        " จาก avg cost";

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("sendBuyByBudget", e);
  }
}




// ----------------------------------------
// Auto Trigger — ตรวจ RSI หุ้นที่ถือทั้งหมด
// แจ้งเตือนเมื่อ RSI < 30 (Oversold)
// ----------------------------------------
function checkRSIOversoldAuto() {
  try {
    const thHoldings = getActiveHoldings(SHEETS.TH_HOLD);
    const usHoldings = getActiveHoldings(SHEETS.US_HOLD);


    const allHoldings = [
      ...thHoldings.map(h => ({ ...h, isTH: true,  yahooSymbol: h.ticker + ".BK" })),
      ...usHoldings.map(h => ({ ...h, isTH: false, yahooSymbol: h.ticker }))
    ];


    if (allHoldings.length === 0) return;


    const oversoldList = [];


    for (const h of allHoldings) {
      const data = _fetchStockData(h.yahooSymbol);
      if (!data) continue;


      if (data.rsi < 30) {
        oversoldList.push({
          ticker: h.ticker,
          isTH:   h.isTH,
          rsi:    data.rsi,
          price:  data.price,
          ma20:   data.ma20,
          ma50:   data.ma50,
          ma200:  data.ma200
        });
      }


      Utilities.sleep(300); // ป้องกัน rate limit
    }


    if (oversoldList.length === 0) return;


    let msg =
      "📉 RSI OVERSOLD ALERT!\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "พบหุ้นที่ RSI ต่ำกว่า 30 (" + oversoldList.length + " ตัว)\n\n";


    for (const r of oversoldList) {
      const cur   = r.isTH ? "฿" : "$";
      const trend = _analyzeTrend(r.price, r.ma20, r.ma50, r.ma200);


      msg +=
        "💡 " + r.ticker + "\n" +
        "  📉 RSI(14)  : " + fmt(r.rsi)        + "\n" +
        "  💰 Price    : " + cur + fmt(r.price) + "\n" +
        "  🎯 Trend    : " + trend.label        + "\n\n";
    }


    msg += "━━━━━━━━━━━━\nพิมพ์ชื่อหุ้นเพื่อดูรายละเอียดเพิ่มเติม";


    sendTelegramSafe(msg);


  } catch (e) {
    logError("checkRSIOversoldAuto", e);
  }
}




// ----------------------------------------
// Setup Trigger — ทุกวัน 09:30 (หลังตลาดเปิด)
// ----------------------------------------
function createRSITrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "checkRSIOversoldAuto")
    .forEach(t => ScriptApp.deleteTrigger(t));


  ScriptApp.newTrigger("checkRSIOversoldAuto")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(30)
    .inTimezone("Asia/Bangkok")
    .create();


  Logger.log("✅ RSI Oversold Trigger: ทุกวัน 09:30");
}


// ========================================
// Risk Assessment Report
// พิมพ์: /risk TICKER
// เช่น: /risk TSM หรือ /risk SCB
// ========================================

function cmdRiskReport(text) {
  try {
    const parts  = text.trim().split(/\s+/);
    const ticker = parts[1] ? parts[1].toUpperCase() : "";

    if (!ticker) {
      sendTelegramSafe(
        "❌ กรุณาระบุชื่อหุ้น\n\n" +
        "✅ วิธีใช้: /risk TICKER\n" +
        "📌 เช่น /risk TSM หรือ /risk SCB"
      );
      return;
    }

    sendTelegramSafe("⏳ กำลังวิเคราะห์ความเสี่ยง " + ticker + "...");

    const isTH  = ticker.endsWith(".BK");
    const yahoo = isTH ? ticker : ticker;
    const data  = _fetchStockData(yahoo);

    if (!data) {
      sendTelegramSafe("❌ ไม่พบข้อมูล: " + ticker);
      return;
    }

    const cur    = isTH ? "฿" : "$";
    const flag   = isTH ? "🇹🇭" : "🇺🇸";
    const price  = data.price;
    const ma20   = data.ma20;
    const ma50   = data.ma50;
    const ma200  = data.ma200;
    const rsi    = data.rsi;
    const vol    = data.volRatio;
    const w52High= data.w52High;
    const w52Low = data.w52Low;

    // ── 1. Trend Status ──
    const trend    = _analyzeTrend(price, ma20, ma50, ma200);
    const underMA20= price < ma20;
    const underMA50= price < ma50;

    let trendStatus = "";
    if (price > ma20 && ma20 > ma50 && ma50 > ma200) {
      trendStatus = "✅ ขาขึ้นแข็งแกร่ง (Strong Uptrend)";
    } else if (price > ma50 && ma50 > ma200) {
      trendStatus = "🟡 ขาขึ้นระยะกลาง";
    } else if (price < ma20 && price > ma50) {
      trendStatus = "🟡 ขาขึ้นระยะกลาง (กำลังพักตัว)";
    } else if (price < ma50 && ma50 > ma200) {
      trendStatus = "🔴 อ่อนแอระยะสั้น (ใต้ MA50)";
    } else {
      trendStatus = "🔴 ขาลง (Downtrend)";
    }

    const techWarning = underMA20
      ? "⚠️ ราคาอยู่ใต้ MA20 — แรงขายระยะสั้นกำลังทดสอบ"
      : "✅ ราคายืนเหนือ MA20";

    const volWarning = vol > 1.5
      ? "⚠️ Volume สูงกว่าเฉลี่ย ×" + fmt(vol) + " — ราคาเหวี่ยงแรง"
      : "✅ Volume ปกติ ×" + fmt(vol);

    // ── 2. Support & Stop Loss ──
    // Support = MA50 (จุดเปลี่ยนเทรนด์)
    const support    = ma50;
    const distSupport= support > 0
      ? ((price - support) / price) * 100 : 0;

    // Stop Loss = MA50 (หลุดแล้วเทรนด์เปลี่ยน)
    const stopLevel  = ma50;

    // Upside Target = MA200 ถ้าราคาต่ำกว่า, หรือ 52W High
    const upTarget   = price < ma200 ? ma200 : w52High;
    const upside     = upTarget > 0
      ? ((upTarget - price) / price) * 100 : 0;
    const downside   = stopLevel > 0
      ? ((price - stopLevel) / price) * 100 : 0;

    // Risk/Reward
    const rr = downside > 0 ? upside / downside : 0;

    // ── 3. RSI Zone ──
    const rsiZone    = _getRSIZone(rsi);
    let   rsiWarning = "";
    if (rsi > 70) {
      rsiWarning = "⚠️ RSI " + fmt(rsi) + " — Overbought อาจพักตัว";
    } else if (rsi < 30) {
      rsiWarning = "💡 RSI " + fmt(rsi) + " — Oversold อาจดีดกลับ";
    } else {
      rsiWarning = "✅ RSI " + fmt(rsi) + " — Neutral";
    }

    // ── 4. Scenario Analysis ──
    const caseA = upTarget;
    const caseB = stopLevel;

    // ── 5. Actionable Status ──
    let actionStatus = "";
    let strategicMove = "";

    if (price > ma20 && ma20 > ma50 && rsi < 70) {
      actionStatus  = "🟢 [BUY SIGNAL] สัญญาณซื้อ";
      strategicMove = "ราคายืนเหนือ MA20 และ MA50 แนวโน้มดี\n" +
        "พิจารณาเข้าซื้อได้ตาม DCA Plan";
    } else if (underMA20 && price > ma50 && rsi < 65) {
      actionStatus  = "👀 [WATCHLIST] รอยืนยัน";
      strategicMove = "รอราคากลับขึ้นเหนือ MA20 (" + cur + fmt(ma20) + ")\n" +
        "หากยืนได้และมี Bullish Reversal ค่อยเข้าซื้อ";
    } else if (underMA50) {
      actionStatus  = "⚠️ [CAUTION] ระวัง";
      strategicMove = "ราคาหลุด MA50 ความเสี่ยงสูง\n" +
        "ยังไม่แนะนำให้เข้าซื้อเพิ่ม รอราคากลับมายืนเหนือ MA50";
    } else if (rsi > 70) {
      actionStatus  = "⏳ [WAIT] รอพักตัว";
      strategicMove = "ราคาวิ่งมาแล้วมาก RSI สูง\n" +
        "รอ RSI ลดลงต่ำกว่า 60 ก่อนค่อยเข้า";
    } else {
      actionStatus  = "🟡 [NEUTRAL] เฝ้าระวัง";
      strategicMove = "แนวโน้มไม่ชัดเจน รอสัญญาณเพิ่มเติม";
    }

    const riskAlert = underMA20
      ? "🚨 ห้ามเทหมดหน้าตัก ราคายังอยู่ใต้ MA20\n" +
        "   ความเสี่ยงลงทดสอบ MA50 ที่ " + cur + fmt(ma50) + " ยังมีอยู่"
      : "✅ ราคายืนเหนือเส้นค่าเฉลี่ยหลัก ความเสี่ยงต่ำ";

    // ── Build Message ──
    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy"
    );

    const msg =
      "⚠️ RISK ASSESSMENT REPORT\n" +
      ticker + " " + flag + "  |  " + today + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "1️⃣ Market Context & Trend\n\n" +
      "📊 Trend Status   : " + trendStatus  + "\n" +
      "📉 Tech Warning   : " + techWarning  + "\n" +
      "📦 Volatility     : " + volWarning   + "\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "2️⃣ Tactical Risk Metrics\n\n" +
      "📏 Distance to Support\n" +
      "   " + cur + fmt(support) + " (MA50) — " +
        fmt(Math.abs(distSupport)) + "% จากราคาปัจจุบัน\n" +
      "   หากหลุดจุดนี้ ความเสี่ยงขาลงเพิ่มขึ้น\n\n" +
      "🛑 Stop Loss Level : " + cur + fmt(stopLevel) + "\n" +
      "   (MA50 — จุดเปลี่ยนเทรนด์)\n\n" +
      "📉 RSI Zone        : " + rsiWarning  + "\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "3️⃣ Scenario Analysis\n\n" +
      "✅ Case A (Best)\n" +
      "   ราคายืนเหนือ " + cur + fmt(stopLevel) + " ได้\n" +
      "   🎯 Upside Target : " + cur + fmt(caseA) +
        " (+" + fmt(upside) + "%)\n\n" +
      "❌ Case B (Worst)\n" +
      "   หลุด " + cur + fmt(caseB) + " ไม่มี Bounce\n" +
      "   📉 พิจารณาตัดขาดทุนทันที\n\n" +
      "⚖️ Risk/Reward Ratio : 1 : " + fmt(rr) + "\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "4️⃣ Actionable Advice\n\n" +
      "📌 Status         : " + actionStatus  + "\n\n" +
      "🎯 Strategic Move :\n" +
      "   " + strategicMove + "\n\n" +
      "🚨 Risk Alert     :\n" +
      "   " + riskAlert;

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("cmdRiskReport", e);
  }
}









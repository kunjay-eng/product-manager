// ========================================
// dca_analysis.gs
// ระบบวิเคราะห์ DCA
// 1) รวมเข้ากับ stockinfo (พิมพ์ชื่อหุ้น)
// 2) /dcareport — วิเคราะห์ทุกตัวที่มี Target
// ========================================


// ----------------------------------------
// ดึง Target Allocation จาก Holdings col P
// ----------------------------------------
function _getDCATargets() {
  const targets = {}; // { ticker: { target, isTH } }


  [
    { sheet: SHEETS.US_HOLD, isTH: false },
    { sheet: SHEETS.TH_HOLD, isTH: true  }
  ].forEach(({ sheet: sheetName, isTH }) => {
    const sheet   = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.HOLD) return;


    const numRows = lastRow - START_ROW.HOLD + 1;
    const rows    = sheet.getRange(
      START_ROW.HOLD, 1, numRows, 16
    ).getValues();


    rows.forEach(row => {
      const ticker = row[HOLD_COL.TICKER - 1];
      const target = Number(row[HOLD_COL.DCA_TARGET - 1]) || 0;
      if (!ticker || target <= 0) return;


      targets[String(ticker).trim().toUpperCase()] = { target, isTH };
    });
  });


  return targets;
}


// ----------------------------------------
// คำนวณ Allocation ปัจจุบัน ของ ticker หนึ่งตัว
// เทียบกับมูลค่าพอร์ตรวม (รวม TH+US+Fund)
// ----------------------------------------
function _getCurrentAllocation(ticker, isTH) {
  const d = collectAllData();
  const holdings = isTH ? d.thHoldings : d.usHoldings;
  const h = holdings.find(x => x.ticker === ticker);
  if (!h) return { valueNow: 0, allocPct: 0 };


  const valueNowTHB = isTH ? h.valueNow : h.valueNow * d.fxRate;
  const allocPct = d.totalPortTHB > 0
    ? (valueNowTHB / d.totalPortTHB) * 100 : 0;


  return { valueNow: h.valueNow, allocPct };
}


// ----------------------------------------
// คำนวณ Support levels จาก MA50/MA200
// Support1 = MA50, Support2 = MA200
// ----------------------------------------
function _calcSupportLevels(ma50, ma200) {
  return {
    support1: ma50,
    support2: ma200
  };
}








/**
 * ====================================================
 *
 * ====================================================
 *
 
// ----------------------------------------
// วิเคราะห์ DCA Signal ของหุ้นตัวเดียว
// ตามกฎจากเอกสาร
// ----------------------------------------
function _analyzeDCASignal(params) {
  const { allocPct, target, price, ma50, rsi } = params;


  const distFromMA50 = ma50 > 0
    ? ((price - ma50) / ma50) * 100 : 0;


  const reasons = [];
  let signal = "";
  let emoji  = "";


  // 🚨 หยุดเติมทันที (เงื่อนไขฉุกเฉิน)
  if (allocPct > target + 20) {
    signal = "PAUSE"; emoji = "🔴";
    reasons.push("🚨 Allocation เกินเป้าหมายมาก (" + fmt(allocPct) + "% vs " + fmt(target) + "%)");
  } else if (rsi > 75) {
    signal = "PAUSE"; emoji = "🔴";
    reasons.push("🚨 RSI > 75 (" + fmt(rsi) + ") — Overbought รุนแรง");
  } else if (distFromMA50 > 15) {
    signal = "PAUSE"; emoji = "🔴";
    reasons.push("🚨 ราคาวิ่งห่าง MA50 เกิน 15% (+" + fmt(distFromMA50) + "%)");
  }
  // 🔴 PAUSE DCA — เงื่อนไขปกติ
  else if (allocPct > target) {
    signal = "PAUSE"; emoji = "🔴";
    reasons.push("❌ Allocation สูงกว่าเป้าหมาย (" + fmt(allocPct) + "% > " + fmt(target) + "%)");
    if (rsi > 70) reasons.push("❌ RSI สูงเกิน 70 (" + fmt(rsi) + ")");
  } else if (rsi > 70) {
    signal = "PAUSE"; emoji = "🔴";
    reasons.push("❌ RSI > 70 (" + fmt(rsi) + ") — Overbought");
  } else if (distFromMA50 > 15) {
    signal = "PAUSE"; emoji = "🔴";
    reasons.push("❌ ราคาห่าง MA50 เกิน 15%");
  }
  // 🟡 WAIT
  else if (allocPct < target && rsi >= 60 && rsi <= 70) {
    signal = "WAIT"; emoji = "🟡";
    reasons.push("⚠️ Allocation ต่ำกว่าเป้า แต่ RSI อยู่ช่วง 60-70 (" + fmt(rsi) + ")");
  } else if (Math.abs(distFromMA50) > 8) {
    signal = "WAIT"; emoji = "🟡";
    reasons.push("⚠️ ราคาห่าง MA50 มากกว่า 8% (" + fmt(distFromMA50) + "%)");
  }
  // 🟢 DCA NOW
  else if (allocPct < target && price > ma50 && rsi < 60) {
    signal = "BUY"; emoji = "🟢";
    reasons.push("✅ Allocation ต่ำกว่าเป้าหมาย (" + fmt(allocPct) + "% < " + fmt(target) + "%)");
    reasons.push("✅ ราคา > MA50 (แนวโน้มดี)");
    reasons.push("✅ RSI ไม่แพง (" + fmt(rsi) + ")");
  } else {
    signal = "WAIT"; emoji = "🟡";
    reasons.push("⏳ ยังไม่เข้าเงื่อนไขชัดเจน รอสัญญาณเพิ่ม");
  }


  return { signal, emoji, reasons, distFromMA50 };
}
// ----------------------------------------
// รวม DCA Analysis เข้ากับ stockinfo
// เรียกจาก stockinfo.gs
// ----------------------------------------
function _appendDCAAnalysis(ticker, stockData) {
  const targets = _getDCATargets();
  const key     = ticker.replace(".BK", "").toUpperCase();
  const targetInfo = targets[key];


  if (!targetInfo) return ""; // ไม่มี target ตั้งไว้ ไม่ต้องแสดง DCA


  const alloc = _getCurrentAllocation(key, targetInfo.isTH);
  const support = _calcSupportLevels(stockData.ma50, stockData.ma200);


  const analysis = _analyzeDCASignal({
    allocPct: alloc.allocPct,
    target:   targetInfo.target,
    price:    stockData.price,
    ma50:     stockData.ma50,
    rsi:      stockData.rsi
  });


  const cur = targetInfo.isTH ? "฿" : "$";


  let msg =
    "\n━━━━━━━━━━━━\n" +
    "🎯 DCA ANALYSIS\n\n" +
    "📊 Allocation\n" +
    "  ปัจจุบัน : " + fmt(alloc.allocPct) + "%\n" +
    "  เป้าหมาย : " + fmt(targetInfo.target) + "%\n\n" +
    "📉 Support Levels\n" +
    "  Support 1 (MA50)  : " + cur + fmt(support.support1) + "\n" +
    "  Support 2 (MA200) : " + cur + fmt(support.support2) + "\n\n" +
    "📋 เหตุผล\n";


  analysis.reasons.forEach(r => { msg += r + "\n"; });


  msg +=
    "\n" + analysis.emoji + " สรุป: " +
    (analysis.signal === "BUY"   ? "DCA NOW — ซื้อได้" :
     analysis.signal === "WAIT"  ? "WAIT — รอจังหวะ" :
                                    "PAUSE DCA — หยุดเติม");


  return msg;
}


// ----------------------------------------
// /dcareport — วิเคราะห์ DCA ทุกตัวที่ตั้ง Target ไว้
// ----------------------------------------
function sendDCAReport() {
  sendTelegramSafe("⏳ กำลังวิเคราะห์ DCA ทุกตัว...");
  try {
    const targets = _getDCATargets();
    const tickers = Object.keys(targets);


    if (tickers.length === 0) {
      sendTelegramSafe(
        "📭 ยังไม่ได้ตั้ง Target Allocation\n\n" +
        "💡 กรอกเปอร์เซ็นต์เป้าหมายใน Holdings col P\n" +
        "(💼 Holdings หรือ 🇹🇭  💼 Holdings)"
      );
      return;
    }


    const results = [];


    for (const ticker of tickers) {
      const info = targets[ticker];
      const yahooSymbol = info.isTH ? ticker + ".BK" : ticker;
      const stockData = _fetchStockData(yahooSymbol);
      if (!stockData) continue;


      const alloc = _getCurrentAllocation(ticker, info.isTH);
      const analysis = _analyzeDCASignal({
        allocPct: alloc.allocPct,
        target:   info.target,
        price:    stockData.price,
        ma50:     stockData.ma50,
        rsi:      stockData.rsi
      });


      results.push({
        ticker, isTH: info.isTH,
        target: info.target,
        allocPct: alloc.allocPct,
        price: stockData.price,
        signal: analysis.signal,
        emoji: analysis.emoji
      });


      Utilities.sleep(300);
    }


    if (results.length === 0) {
      sendTelegramSafe("❌ ไม่สามารถดึงข้อมูลหุ้นได้");
      return;
    }


    let msg =
      "📊 DCA REPORT\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";


    results.forEach(r => {
      const cur = r.isTH ? "฿" : "$";
      msg +=
        r.emoji + " " + r.ticker + "\n" +
        "  Allocation : " + fmt(r.allocPct) + "% / เป้า " + fmt(r.target) + "%\n" +
        "  Price      : " + cur + fmt(r.price) + "\n" +
        "  สถานะ      : " +
          (r.signal === "BUY"   ? "🟢 BUY" :
           r.signal === "WAIT"  ? "🟡 WAIT" : "🔴 PAUSE") +
        "\n\n";
    });


    // Portfolio Recommendation
    const buyList   = results.filter(r => r.signal === "BUY");
    const pauseList = results.filter(r => r.signal === "PAUSE");


    msg += "━━━━━━━━━━━━\n📋 Portfolio Recommendation\n\n";


    if (buyList.length > 0) {
      buyList.forEach(r => {
        msg += "➡️ เติม " + r.ticker + "\n";
      });
    }
    if (pauseList.length > 0) {
      pauseList.forEach(r => {
        msg += "🚫 หยุดเติม " + r.ticker + "\n";
      });
    }
    if (buyList.length === 0 && pauseList.length === 0) {
      msg += "⏳ ทุกตัวอยู่ในสถานะ WAIT — รอจังหวะ\n";
    }


    sendTelegramSafe(msg);


  } catch (e) {
    sendTelegramError("sendDCAReport", e);
  }
}


*/


// ----------------------------------------
// คำนวณ DCA Score (0-10 คะแนน)
// ตามเกณฑ์ใหม่
// ----------------------------------------
function _calcDCAScore(params) {
  const { allocPct, target, price, ma50, ma200, rsi, support } = params;


  let score = 0;
  const breakdown = [];


  // 1) Allocation ต่ำกว่าเป้าหมาย +3
  if (allocPct < target) {
    score += 3;
    breakdown.push("✅ Allocation ต่ำกว่าเป้าหมาย (+3)");
  } else {
    breakdown.push("❌ Allocation ถึง/เกินเป้าหมาย (+0)");
  }


  // 2) Price > MA50 +2
  if (price > ma50) {
    score += 2;
    breakdown.push("✅ Price > MA50 (+2)");
  } else {
    breakdown.push("❌ Price < MA50 (+0)");
  }


  // 3) Price > MA200 +1
  if (price > ma200) {
    score += 1;
    breakdown.push("✅ Price > MA200 (+1)");
  } else {
    breakdown.push("❌ Price < MA200 (+0)");
  }


  // 4) RSI < 60 +2
  if (rsi < 60) {
    score += 2;
    breakdown.push("✅ RSI < 60 (+2)");
  } else {
    breakdown.push("❌ RSI ≥ 60 (+0)");
  }


  // 5) ใกล้ MA50 (≤5%) +1
  const distMA50 = ma50 > 0 ? Math.abs((price - ma50) / ma50) * 100 : 999;
  if (distMA50 <= 5) {
    score += 1;
    breakdown.push("✅ ใกล้ MA50 ≤5% (+1)");
  } else {
    breakdown.push("❌ ไม่ใกล้ MA50 (+0)");
  }


  // 6) ใกล้แนวรับ (≤5%) +1
  const distSupport = support > 0 ? ((price - support) / support) * 100 : 999;
  const nearSupport = distSupport <= 5;
  if (nearSupport) {
    score += 1;
    breakdown.push("✅ ใกล้แนวรับ ≤5% (+1)");
  } else {
    breakdown.push("❌ ไม่ใกล้แนวรับ (+0)");
  }


  // สรุประดับ
  let signal, emoji, label;
  if (score >= 8) {
    signal = "BUY"; emoji = "🟢"; label = "DCA NOW";
  } else if (score >= 5) {
    signal = "WAIT"; emoji = "🟡"; label = "WAIT FOR DIP";
  } else {
    signal = "PAUSE"; emoji = "🔴"; label = "PAUSE DCA";
  }


  return {
    score, breakdown, signal, emoji, label,
    distMA50, distSupport, nearSupport
  };
}


// ----------------------------------------
// RSI Zone (ใช้แสดงเพิ่มเติม)
// ----------------------------------------
function _getRSIZone(rsi) {
  if (rsi > 80)               return { emoji: "🔴", label: "หยุดเติม รอพักฐาน" };
  if (rsi > 70)                return { emoji: "🔴", label: "ชะลอการ DCA" };
  if (rsi >= 60)                return { emoji: "🟡", label: "ซื้อได้แต่ไม่ต้องเร่ง" };
  if (rsi >= 40)                return { emoji: "🟢", label: "สะสมได้" };
  return { emoji: "🟢", label: "น่าสะสมมาก" };
}


// ----------------------------------------
// รวม DCA Analysis เข้ากับ stockinfo
// ----------------------------------------


function _appendDCAAnalysis(ticker, stockData) {
  const targets = _getDCATargets();
  const key     = ticker.replace(".BK", "").toUpperCase();
  const targetInfo = targets[key];


  if (!targetInfo) return "";


  const alloc   = _getCurrentAllocation(key, targetInfo.isTH);
  const support = _calcSupportLevels(stockData.ma50, stockData.ma200);


  const scoreParams = {
    allocPct: alloc.allocPct,
    target:   targetInfo.target,
    price:    stockData.price,
    ma50:     stockData.ma50,
    ma200:    stockData.ma200,
    rsi:      stockData.rsi,
    support:  support.support1
  };


  const result = _calcDCAScore(scoreParams);
  const rsiZone = _getRSIZone(stockData.rsi);
  const cur = targetInfo.isTH ? "฿" : "$";


  const explanation = _explainDCAScore(result, scoreParams);


  // ✅ ดึงงบประมาณจาก Portfolio sheet (ถ้ามี)
  const budget = _getDCABudget(key, targetInfo.isTH);


  let msg =
    "\n━━━━━━━━━━━━\n" +
    "🎯 DCA ANALYSIS\n\n" +


    "📊 Allocation\n" +
    "  ปัจจุบัน : " + fmt(alloc.allocPct) + "%\n" +
    "  เป้าหมาย : " + fmt(targetInfo.target) + "%\n\n" +


    "📉 Support & Distance\n" +
    "  Support (MA50)  : " + cur + fmt(support.support1) + "\n" +
    "  Distance        : " + fmt(result.distSupport) + "%\n" +
    "  " + (result.nearSupport ? "✅ Near Support" : "❌ Not Near Support") + "\n\n" +


    "📉 RSI Zone\n" +
    "  RSI(14) : " + fmt(stockData.rsi) + "\n" +
    "  " + rsiZone.emoji + " " + rsiZone.label + "\n\n" +


    "🧮 DCA SCORE: " + result.score + "/10\n";


  result.breakdown.forEach(b => { msg += "  " + b + "\n"; });


  msg +=
    "\n" + result.emoji + " สรุป: " + result.label + "\n\n" +
    "━━━━━━━━━━━━\n" +
    explanation;


  // ✅ เพิ่มแผนแบ่งไม้ ถ้ามีงบประมาณตั้งไว้
  if (budget > 0) {
    const plan = _buildLotPlan({
      score: result.score,
      price: stockData.price,
      ma50:  stockData.ma50,
      ma200: stockData.ma200,
      budget,
      currency: cur
    });


    msg += "\n\n━━━━━━━━━━━━\n" + _formatLotPlan(plan, cur);
  } else {
    msg += "\n\n💡 ตั้งงบประมาณใน Portfolio sheet col M เพื่อดูแผนแบ่งไม้";
  }


  return msg;
}


// ----------------------------------------
// ดึงงบประมาณจาก Portfolio sheet col M
// ----------------------------------------
function _getDCABudget(ticker, isTH) {
  const sheetName = isTH ? SHEETS.TH_HOLD : SHEETS.US_HOLD;
  const sheet     = getSheet(sheetName);
  const lastRow   = sheet.getLastRow();
  if (lastRow < START_ROW.HOLD) return 0;


  const numRows = lastRow - START_ROW.HOLD + 1;
  const rows    = sheet.getRange(
    START_ROW.HOLD, 1, numRows, 17
  ).getValues();


  for (const row of rows) {
    const rowTicker = String(row[HOLD_COL.TICKER - 1] || "").trim().toUpperCase();
    if (rowTicker === ticker) {
      return Number(row[HOLD_COL.DCA_BUDGET - 1]) || 0;
    }
  }


  return 0;
}




// ----------------------------------------
// Entry Analysis สำหรับหุ้นใหม่ที่ยังไม่ถือ
// (ไม่มี Target/Allocation เพราะยังไม่มีในพอร์ต)
// ----------------------------------------
function _appendEntryAnalysis(ticker, stockData) {
  const support = _calcSupportLevels(stockData.ma50, stockData.ma200);
  const rsiZone = _getRSIZone(stockData.rsi);


  const distMA50 = stockData.ma50 > 0
    ? ((stockData.price - stockData.ma50) / stockData.ma50) * 100 : 0;
  const distSupport = support.support1 > 0
    ? ((stockData.price - support.support1) / support.support1) * 100 : 0;
  const nearSupport = Math.abs(distSupport) <= 5;


  // ให้คะแนนแบบง่าย (ไม่มี Allocation เพราะยังไม่ถือ)
  // เต็ม 7 คะแนน (ตัด Allocation +3 ออก)
  let score = 0;
  const breakdown = [];


  if (stockData.price > stockData.ma50) {
    score += 2;
    breakdown.push("✅ Price > MA50 (+2)");
  } else {
    breakdown.push("❌ Price < MA50 (+0)");
  }


  if (stockData.price > stockData.ma200) {
    score += 1;
    breakdown.push("✅ Price > MA200 (+1)");
  } else {
    breakdown.push("❌ Price < MA200 (+0)");
  }


  if (stockData.rsi < 60) {
    score += 2;
    breakdown.push("✅ RSI < 60 (+2)");
  } else {
    breakdown.push("❌ RSI ≥ 60 (+0)");
  }


  if (Math.abs(distMA50) <= 5) {
    score += 1;
    breakdown.push("✅ ใกล้ MA50 ≤5% (+1)");
  } else {
    breakdown.push("❌ ไม่ใกล้ MA50 (+0)");
  }


  if (nearSupport) {
    score += 1;
    breakdown.push("✅ ใกล้แนวรับ ≤5% (+1)");
  } else {
    breakdown.push("❌ ไม่ใกล้แนวรับ (+0)");
  }


  let signal, emoji, label;
  if (score >= 6) {
    signal = "BUY"; emoji = "🟢"; label = "น่าเข้าซื้อ";
  } else if (score >= 3) {
    signal = "WAIT"; emoji = "🟡"; label = "รอจังหวะดีกว่านี้";
  } else {
    signal = "AVOID"; emoji = "🔴"; label = "ยังไม่น่าเข้า";
  }


  const cur = stockData.isTH ? "฿" : "$";


  let msg =
    "\n━━━━━━━━━━━━\n" +
    "🆕 ENTRY ANALYSIS (หุ้นใหม่ ยังไม่ถือ)\n\n" +


    "📉 Support & Distance\n" +
    "  Support (MA50)  : " + cur + fmt(support.support1) + "\n" +
    "  Distance        : " + fmt(distSupport) + "%\n" +
    "  " + (nearSupport ? "✅ Near Support" : "❌ Not Near Support") + "\n\n" +


    "📉 RSI Zone\n" +
    "  RSI(14) : " + fmt(stockData.rsi) + "\n" +
    "  " + rsiZone.emoji + " " + rsiZone.label + "\n\n" +


    "🧮 ENTRY SCORE: " + score + "/7\n";


  breakdown.forEach(b => { msg += "  " + b + "\n"; });


  msg +=
    "\n" + emoji + " สรุป: " + label + "\n\n" +
    "💡 หมายเหตุ: หุ้นนี้ยังไม่อยู่ในพอร์ต\n" +
    "หากต้องการติดตาม DCA Score แบบเต็ม กรุณาซื้อแล้วตั้ง\n" +
    "Target Allocation ใน Holdings col P ก่อนครับ";


  return msg;
}




// ----------------------------------------
// sendDCAReport()
//
// ----------------------------------------


function sendDCAReport() {
  sendTelegramSafe("⏳ กำลังวิเคราะห์ DCA ทุกตัว...");
  try {
    const targets = _getDCATargets();
    const tickers = Object.keys(targets);


    if (tickers.length === 0) {
      sendTelegramSafe(
        "📭 ยังไม่ได้ตั้ง Target Allocation\n\n" +
        "💡 กรอกเปอร์เซ็นต์เป้าหมายใน Holdings col P"
      );
      return;
    }


    const results = [];


    for (const ticker of tickers) {
      const info = targets[ticker];
      const yahooSymbol = info.isTH ? ticker + ".BK" : ticker;
      const stockData = _fetchStockData(yahooSymbol);
      if (!stockData) continue;


      const alloc   = _getCurrentAllocation(ticker, info.isTH);
      const support = _calcSupportLevels(stockData.ma50, stockData.ma200);


      const scoreParams = {
        allocPct: alloc.allocPct,
        target:   info.target,
        price:    stockData.price,
        ma50:     stockData.ma50,
        ma200:    stockData.ma200,
        rsi:      stockData.rsi,
        support:  support.support1
      };


      const result = _calcDCAScore(scoreParams);


      // ✅ ดึงสาเหตุหลักข้อเดียว (สั้นๆ) มาแสดงใน list
      const explanation = _explainDCAScore(result, scoreParams);
      const mainReason  = explanation.split("\n")[0]
        .replace("🔍 สาเหตุหลักที่คะแนนไม่เต็ม: ", "")
        .replace("✨ ", "");


      results.push({
        ticker, isTH: info.isTH,
        target: info.target,
        allocPct: alloc.allocPct,
        price: stockData.price,
        score: result.score,
        signal: result.signal,
        emoji: result.emoji,
        label: result.label,
        mainReason
      });


      Utilities.sleep(300);
    }


    if (results.length === 0) {
      sendTelegramSafe("❌ ไม่สามารถดึงข้อมูลหุ้นได้");
      return;
    }


    results.sort((a, b) => b.score - a.score);


    let msg =
      "📊 DCA REPORT\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n";


    results.forEach(r => {
      const cur = r.isTH ? "฿" : "$";
      msg +=
        r.emoji + " " + r.ticker + " — " + r.score + "/10\n" +
        "  Allocation : " + fmt(r.allocPct) + "% / เป้า " + fmt(r.target) + "%\n" +
        "  Price      : " + cur + fmt(r.price) + "\n" +
        "  สถานะ      : " + r.label + "\n" +
        "  เหตุผล     : " + r.mainReason + "\n\n";  // ✅ เพิ่ม
    });


    const buyList   = results.filter(r => r.signal === "BUY");
    const pauseList = results.filter(r => r.signal === "PAUSE");


    msg += "━━━━━━━━━━━━\n📋 Portfolio Recommendation\n\n";


    if (buyList.length > 0) {
      buyList.forEach(r => {
        msg += "🟢 เติม " + r.ticker + " (" + r.score + "/10)\n";
      });
    }
    if (pauseList.length > 0) {
      pauseList.forEach(r => {
        msg += "🔴 หยุดเติม " + r.ticker + " — " + r.mainReason + "\n";
      });
    }
    if (buyList.length === 0 && pauseList.length === 0) {
      msg += "🟡 ทุกตัวอยู่ในสถานะ WAIT — รอจังหวะ\n";
    }


    sendTelegramSafe(msg);


  } catch (e) {
    sendTelegramError("sendDCAReport", e);
  }
}


// ----------------------------------------
// วิเคราะห์ breakdown อัตโนมัติ
// บอกว่าเสียคะแนนจากอะไร และควรทำอะไร
// ----------------------------------------
function _explainDCAScore(result, params) {
  const { allocPct, target, rsi, distMA50, distSupport } = params;
  const lost = []; // จุดที่เสียคะแนน เรียงตามน้ำหนัก


  // เช็คแต่ละข้อที่ไม่ผ่าน พร้อมน้ำหนักคะแนนที่เสีย
  if (allocPct >= target) {
    lost.push({ pts: 3, text: "Allocation ถึง/เกินเป้าหมายแล้ว (สาเหตุหลัก)" });
  }
  if (params.price <= params.ma50) {
    lost.push({ pts: 2, text: "ราคาหลุด MA50 — แนวโน้มระยะสั้นอ่อนแอ" });
  }
  if (rsi >= 60) {
    lost.push({ pts: 2, text: "RSI สูงเกิน 60 — ราคาเริ่มแพง" });
  }
  if (params.price <= params.ma200) {
    lost.push({ pts: 1, text: "ราคาหลุด MA200 — แนวโน้มระยะยาวอ่อนแอ" });
  }
  if (distMA50 > 5) {
    lost.push({ pts: 1, text: "ราคาห่าง MA50 เกิน 5%" });
  }
  if (distSupport > 5) {
    lost.push({ pts: 1, text: "ราคายังไม่ใกล้แนวรับ" });
  }


  // เรียงจากเสียคะแนนเยอะสุดก่อน
  lost.sort((a, b) => b.pts - a.pts);


  // สร้างคำอธิบาย
  let explain = "";


  if (lost.length === 0) {
    explain = "✨ ผ่านครบทุกเงื่อนไข เหมาะกับการสะสมตอนนี้มากที่สุด";
  } else {
    const top = lost[0];
    explain = "🔍 สาเหตุหลักที่คะแนนไม่เต็ม: " + top.text;


    if (lost.length > 1) {
      explain += "\n   (และอีก " + (lost.length - 1) + " ข้อรอง: " +
        lost.slice(1).map(l => l.text).join(", ") + ")";
    }
  }


  // คำแนะนำการกระทำ ตามสาเหตุหลัก
  let action = "";
  if (lost.length === 0) {
    action = "💡 แนะนำ: เติมตามแผนได้เลย";
  } else if (lost[0].text.includes("Allocation")) {
    action = "💡 แนะนำ: รอสัดส่วนพอร์ตลดลงก่อน หรือ rebalance ตัวอื่น";
  } else if (lost[0].text.includes("RSI")) {
    action = "💡 แนะนำ: รอ RSI ย่อตัวลงมาก่อนค่อยเติม";
  } else if (lost[0].text.includes("MA50") || lost[0].text.includes("MA200")) {
    action = "💡 แนะนำ: รอราคากลับขึ้นเหนือเส้นค่าเฉลี่ยก่อน";
  } else {
    action = "💡 แนะนำ: รอราคาย่อใกล้แนวรับมากขึ้น";
  }


  return explain + "\n" + action;
}




// ----------------------------------------
// กำหนดจำนวนไม้ตาม DCA Score
// ----------------------------------------
function _getLotCount(score) {
  if (score >= 8) return 2;
  if (score >= 5) return 3;
  return 4;
}


// ----------------------------------------
// Pyramid weight % ตามจำนวนไม้
// ----------------------------------------
function _getPyramidWeights(lotCount) {
  const weights = {
    2: [40, 60],
    3: [20, 30, 50],
    4: [15, 20, 25, 40]
  };
  return weights[lotCount] || weights[3];
}


// ----------------------------------------
// คำนวณราคาแต่ละไม้
// ผสมระหว่าง % step กับ Support (MA50/MA200)
// ----------------------------------------
function _calcLotPrices(price, ma50, ma200, lotCount) {
  const prices = [];


  // ไม้แรก = ราคาปัจจุบัน หรือ MA50 (เลือกตัวที่ต่ำกว่า ถ้าใกล้กัน)
  prices.push(price);


  if (lotCount === 2) {
    // ไม้ 2 = MA50 ถ้าต่ำกว่าไม้แรกมากพอ ไม่งั้นใช้ -8%
    const lot2 = ma50 < price * 0.97 ? ma50 : price * 0.92;
    prices.push(lot2);
  }


  else if (lotCount === 3) {
    const lot2 = ma50 < price * 0.97 ? ma50 : price * 0.93;
    const lot3 = ma200 < lot2 * 0.95 ? ma200 : lot2 * 0.90;
    prices.push(lot2, lot3);
  }


  else if (lotCount === 4) {
    const lot2 = ma50 < price * 0.97 ? ma50 : price * 0.93;
    const lot3 = ma200 < lot2 * 0.95 ? ma200 : lot2 * 0.90;
    const lot4 = lot3 * 0.85; // -15% จากไม้ 3 (กรณีร่วงหนัก)
    prices.push(lot2, lot3, lot4);
  }


  return prices;
}


// ----------------------------------------
// สร้างแผนแบ่งไม้ทั้งหมด
// ----------------------------------------
function _buildLotPlan(params) {
  const { score, price, ma50, ma200, budget, currency } = params;


  const lotCount = _getLotCount(score);
  const weights  = _getPyramidWeights(lotCount);
  const prices   = _calcLotPrices(price, ma50, ma200, lotCount);


  const lots = [];
  for (let i = 0; i < lotCount; i++) {
    const lotBudget = budget * (weights[i] / 100);
    const lotShares = prices[i] > 0 ? lotBudget / prices[i] : 0;


    lots.push({
      no:      i + 1,
      price:   prices[i],
      weight:  weights[i],
      budget:  lotBudget,
      shares:  lotShares,
      trigger: i === 0 ? "ตอนนี้" :
               (prices[i] === ma50  ? "MA50"  :
                prices[i] === ma200 ? "MA200" : "Price Step")
    });
  }


  return { lotCount, lots };
}


// ----------------------------------------
// แสดงแผนแบ่งไม้เป็นข้อความ
// ----------------------------------------
function _formatLotPlan(plan, currency) {
  let msg =
    "📦 แผนแบ่งไม้ (" + plan.lotCount + " ไม้)\n\n";


  plan.lots.forEach(lot => {
    msg +=
      "ไม้ " + lot.no + " — " + fmt(lot.weight) + "% ของงบ\n" +
      "  💰 ราคา    : " + currency + fmt(lot.price)  + "\n" +
      "  💵 งบ      : " + currency + fmt(lot.budget)  + "\n" +
      "  📦 ซื้อได้  : " + fmt(lot.shares) + " หุ้น\n" +
      "  🎯 อ้างอิง : " + lot.trigger + "\n\n";
  });


  return msg;
}


// ----------------------------------------
// Manual DCA พร้อมคำนวณต้นทุนเฉลี่ยใหม่
// รูปแบบ: TICKER PRICE BUDGET
// เช่น: VOO 650 1000
// ----------------------------------------
function _appendDCAAnalysisManual(ticker, stockData, manualPrice, manualBudget) {
  const cur     = stockData.isTH ? "฿" : "$";
  const support = _calcSupportLevels(stockData.ma50, stockData.ma200);
  const rsiZone = _getRSIZone(stockData.rsi);
  const key     = ticker.replace(".BK", "").toUpperCase();

  // ✅ ดึงข้อมูลที่ถืออยู่ปัจจุบัน
  const holding = _getHoldingInfo(key);

  // ----------------------------------------
  // กรณีที่ 1: มีหุ้นในพอร์ตอยู่แล้ว
  // ----------------------------------------
  if (holding && holding.sharesRemain > 0) {
    return _buildManualWithHolding(
      ticker, stockData, manualPrice, manualBudget,
      holding, cur, support, rsiZone
    );
  }

  // ----------------------------------------
  // กรณีที่ 2: ยังไม่มีหุ้น ซื้อใหม่
  // ----------------------------------------
  return _buildManualNewEntry(
    ticker, stockData, manualPrice, manualBudget,
    cur, support, rsiZone
  );
}

// ----------------------------------------
// ดึงข้อมูลหุ้นที่ถืออยู่จาก Holdings
// ----------------------------------------
function _getHoldingInfo(ticker) {
  const sheets = [
    { name: SHEETS.US_HOLD, isTH: false },
    { name: SHEETS.TH_HOLD, isTH: true  }
  ];

  for (const { name, isTH } of sheets) {
    const sheet   = getSheet(name);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.HOLD) continue;

    const numRows = lastRow - START_ROW.HOLD + 1;
    const rows    = sheet.getRange(
      START_ROW.HOLD, 1, numRows, 17
    ).getValues();

    for (const row of rows) {
      const t = String(row[HOLD_COL.TICKER - 1] || "").trim().toUpperCase();
      if (t !== ticker) continue;

      const sharesRemain = Number(row[HOLD_COL.SHARES_REMAIN - 1]) || 0;
      if (sharesRemain <= 0) return null;

      return {
        ticker,
        isTH,
        sharesRemain,
        avgCost:   Number(row[HOLD_COL.AVG_COST   - 1]) || 0,
        totalCost: Number(row[HOLD_COL.TOTAL_COST - 1]) || 0,
        priceNow:  Number(row[HOLD_COL.PRICE_NOW  - 1]) || 0
      };
    }
  }
  return null;
}

// ----------------------------------------
// กรณี 1: มีหุ้นอยู่แล้ว — คำนวณ avg ใหม่
// ----------------------------------------
function _buildManualWithHolding(
  ticker, stockData, manualPrice, manualBudget,
  holding, cur, support, rsiZone
) {
  const newShares   = manualPrice > 0 ? manualBudget / manualPrice : 0;
  const newTotalCost = holding.totalCost + manualBudget;
  const newTotalShares = holding.sharesRemain + newShares;
  const newAvgCost  = newTotalShares > 0 ? newTotalCost / newTotalShares : 0;
  const currentPL   = holding.priceNow > 0
    ? ((holding.priceNow - holding.avgCost) / holding.avgCost) * 100 : 0;
  const newPL = holding.priceNow > 0
    ? ((holding.priceNow - newAvgCost) / newAvgCost) * 100 : 0;
  const costReduce  = holding.avgCost - newAvgCost;

  // DCA Score (Allocation ไม่รู้ ใช้ 0)
  const scoreParams = {
    allocPct: 0, target: 100,
    price: stockData.price, ma50: stockData.ma50,
    ma200: stockData.ma200, rsi: stockData.rsi,
    support: support.support1
  };
  const result      = _calcDCAScore(scoreParams);
  const explanation = _explainDCAScore(result, scoreParams);

  // แผนแบ่งไม้
  const plan = _buildLotPlan({
    score:    result.score,
    price:    manualPrice,
    ma50:     stockData.ma50,
    ma200:    stockData.ma200,
    budget:   manualBudget,
    currency: cur
  });

  let msg =
    "\n━━━━━━━━━━━━\n" +
    "🔄 DCA ANALYSIS — เพิ่มหุ้นที่ถืออยู่\n\n" +

    "📦 สถานะปัจจุบัน\n" +
    "  ถือหุ้น    : " + fmt(holding.sharesRemain) + " หุ้น\n" +
    "  ต้นทุน avg : " + cur + fmt(holding.avgCost)   + "\n" +
    "  ราคาตลาด  : " + cur + fmt(holding.priceNow)  + "\n" +
    "  P/L        : " + signPct(currentPL)             + "\n\n" +

    "➕ แผนซื้อเพิ่ม\n" +
    "  ราคาที่จะซื้อ : " + cur + fmt(manualPrice)  + "\n" +
    "  งบประมาณ    : " + cur + fmt(manualBudget)   + "\n" +
    "  จำนวนที่ได้   : " + fmt(newShares)          + " หุ้น\n\n" +

    "📊 ต้นทุนเฉลี่ยใหม่\n" +
    "  avg เดิม  : " + cur + fmt(holding.avgCost) + "\n" +
    "  avg ใหม่  : " + cur + fmt(newAvgCost)      + "\n" +
    "  " + (costReduce > 0 ? "✂️ ลดต้นทุน" : "📈 เพิ่มต้นทุน") +
      " : " + cur + fmt(Math.abs(costReduce)) + "/หุ้น\n" +
    "  P/L ใหม่  : " + signPct(newPL)             + "\n\n" +

    "📉 RSI Zone\n" +
    "  RSI(14) : " + fmt(stockData.rsi) + "\n" +
    "  " + rsiZone.emoji + " " + rsiZone.label + "\n\n" +

    "🧮 DCA SCORE: " + result.score + "/10\n";

  result.breakdown.forEach(b => { msg += "  " + b + "\n"; });

  msg +=
    "\n" + result.emoji + " สรุป: " + result.label + "\n\n" +
    "━━━━━━━━━━━━\n" +
    explanation + "\n\n" +
    "━━━━━━━━━━━━\n" +
    _formatLotPlan(plan, cur);

  return msg;
}

// ----------------------------------------
// กรณี 2: ซื้อใหม่ ยังไม่เคยถือ
// ----------------------------------------
function _buildManualNewEntry(
  ticker, stockData, manualPrice, manualBudget,
  cur, support, rsiZone
) {
  const newShares = manualPrice > 0 ? manualBudget / manualPrice : 0;

  // DCA Score
  const scoreParams = {
    allocPct: 0, target: 100,
    price: stockData.price, ma50: stockData.ma50,
    ma200: stockData.ma200, rsi: stockData.rsi,
    support: support.support1
  };
  const result      = _calcDCAScore(scoreParams);
  const explanation = _explainDCAScore(result, scoreParams);

  // แผนแบ่งไม้
  const plan = _buildLotPlan({
    score:    result.score,
    price:    manualPrice,
    ma50:     stockData.ma50,
    ma200:    stockData.ma200,
    budget:   manualBudget,
    currency: cur
  });

  let msg =
    "\n━━━━━━━━━━━━\n" +
    "🆕 DCA ANALYSIS — ซื้อใหม่\n\n" +

    "📦 แผนเข้าซื้อครั้งแรก\n" +
    "  ราคาที่จะซื้อ : " + cur + fmt(manualPrice)  + "\n" +
    "  งบประมาณ    : " + cur + fmt(manualBudget)   + "\n" +
    "  จำนวนที่ได้   : " + fmt(newShares)          + " หุ้น\n" +
    "  ต้นทุน avg   : " + cur + fmt(manualPrice)   + "\n\n" +

    "📉 RSI Zone\n" +
    "  RSI(14) : " + fmt(stockData.rsi) + "\n" +
    "  " + rsiZone.emoji + " " + rsiZone.label + "\n\n" +

    "🧮 DCA SCORE: " + result.score + "/10\n";

  result.breakdown.forEach(b => { msg += "  " + b + "\n"; });

  msg +=
    "\n" + result.emoji + " สรุป: " + result.label + "\n\n" +
    "━━━━━━━━━━━━\n" +
    explanation + "\n\n" +
    "━━━━━━━━━━━━\n" +
    _formatLotPlan(plan, cur);

  return msg;
}


// ========================================
// DCA Alert Auto — ทุกวันจันทร์เช้า
// ส่ง DCA Score ทุกตัวที่ตั้ง Target ไว้
// ========================================
function sendDCAAlertAuto() {
  try {
    const targets = _getDCATargets();
    const tickers = Object.keys(targets);
    if (tickers.length === 0) return;

    const buyList  = [];
    const waitList = [];

    for (const ticker of tickers) {
      const info = targets[ticker];
      const yahooSymbol = info.isTH ? ticker + ".BK" : ticker;
      const stockData = _fetchStockData(yahooSymbol);
      if (!stockData) continue;

      const alloc   = _getCurrentAllocation(ticker, info.isTH);
      const support = _calcSupportLevels(stockData.ma50, stockData.ma200);

      const result = _calcDCAScore({
        allocPct: alloc.allocPct,
        target:   info.target,
        price:    stockData.price,
        ma50:     stockData.ma50,
        ma200:    stockData.ma200,
        rsi:      stockData.rsi,
        support:  support.support1
      });

      const cur = info.isTH ? "฿" : "$";

      if (result.signal === "BUY") {
        buyList.push({
          ticker,
          isTH: info.isTH,
          score: result.score,
          price: stockData.price,
          rsi: stockData.rsi,
          cur
        });
      } else if (result.signal === "WAIT") {
        waitList.push({
          ticker,
          score: result.score,
          cur
        });
      }

      Utilities.sleep(300);
    }

    // ส่งเฉพาะถ้ามี BUY signal
    if (buyList.length === 0) return;

    let msg =
      "🎯 DCA ALERT — สัปดาห์นี้\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "🟢 น่าสะสมตอนนี้ (" + buyList.length + " ตัว)\n\n";

    buyList.forEach(r => {
      const flag = r.isTH ? "🇹🇭" : "🇺🇸";
      msg +=
        flag + " " + r.ticker + " — " + r.score + "/10\n" +
        "  ราคา : " + r.cur + fmt(r.price) + "\n" +
        "  RSI  : " + fmt(r.rsi) + "\n\n";
    });

    if (waitList.length > 0) {
      msg += "━━━━━━━━━━━━\n" +
        "🟡 รอจังหวะ (" + waitList.length + " ตัว)\n";
      waitList.forEach(r => {
        msg += "  • " + r.ticker + " (" + r.score + "/10)\n";
      });
    }

    sendTelegramSafe(msg, "DCA_ALERT", "");
  } catch (e) {
    logError("sendDCAAlertAuto", e);
  }
}

function createDCAAlertTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendDCAAlertAuto")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendDCAAlertAuto")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(30)
    .inTimezone("Asia/Bangkok")
    .create();

  Logger.log("✅ DCA Alert Trigger: ทุกวันจันทร์ 08:30");
}

// ========================================
// Backtest อย่างง่าย
// คำนวณผลตอบแทนจากวันที่ซื้อถึงปัจจุบัน
// /backtest VOO หรือ /backtest VOO 2026-01-01
// ========================================
/**
 * ====================================================
 * ### ข้อควรระวัง  ถ้าระบุวันที่เกิน range ที่ดึงมา (1 ปี) จะได้ข้อมูลน้อย เช่น `/backtest MU 2024-01-01` อาจไม่มีข้อมูลเพราะดึงแค่ 365 วันย้อนหลัง ต้องการข้อมูลนานกว่า 1 ปีต้องแก้ `fetchYahooHistory(ticker, 365)` เป็น `fetchYahooHistory(ticker, 730)` หรือมากกว่าครับ
 * ====================================================
 * 
 */

function cmdBacktest(text) {
  try {
    const parts    = text.trim().split(/\s+/);
    const ticker   = parts[1] ? parts[1].toUpperCase() : "";
    const fromDate = parts[2] || null;

    if (!ticker) {
      sendTelegramSafe(
        "❌ กรุณาระบุชื่อหุ้น\n\n" +
        "✅ วิธีใช้:\n" +
        "/backtest TICKER\n" +
        "/backtest TICKER YYYY-MM-DD\n\n" +
        "📌 ตัวอย่าง:\n" +
        "/backtest MU\n" +
        "/backtest MU 2026-01-01"
      );
      return;
    }

    sendTelegramSafe("⏳ กำลังคำนวณ Backtest " + ticker + "...");

    const history = fetchYahooHistory(ticker, 365);
    if (!history || history.length === 0) {
      sendTelegramSafe("❌ ไม่พบข้อมูลย้อนหลัง: " + ticker);
      return;
    }

    let filtered = history;
    if (fromDate) {
      const from = new Date(fromDate);
      filtered = history.filter(d => d.date >= from);
    }

    if (filtered.length < 2) {
      sendTelegramSafe("❌ ข้อมูลน้อยเกินไป ลองเปลี่ยนวันที่");
      return;
    }

    const first   = filtered[0];
    const last    = filtered[filtered.length - 1];
    const current = last.close;
    const buyAt   = first.close;

    // ── Lump Sum ──
    const lsReturn  = ((current - buyAt) / buyAt) * 100;
    const maxClose  = Math.max(...filtered.map(d => d.close));
    const minClose  = Math.min(...filtered.map(d => d.close));

    // ✅ Max Drawdown คำนวณถูกต้อง
    // หา peak ก่อนแล้วคำนวณ drawdown จาก peak นั้น
    let maxDD = 0;
    let peak  = filtered[0].close;
    filtered.forEach(d => {
      if (d.close > peak) peak = d.close;
      const dd = ((d.close - peak) / peak) * 100;
      if (dd < maxDD) maxDD = dd;
    });

    // ✅ CAGR
    const days  = Math.max(1,
      (last.date - first.date) / (1000 * 60 * 60 * 24));
    const years = days / 365;
    const lsCAGR = years > 0
      ? (Math.pow(current / buyAt, 1 / years) - 1) * 100 : 0;

    // ── DCA ทุกสัปดาห์ ──
    const weeklyBudget = 1000;
    let dcaShares = 0, dcaCost = 0;
    filtered.filter((_, i) => i % 5 === 0).forEach(d => {
      dcaShares += weeklyBudget / d.close;
      dcaCost   += weeklyBudget;
    });
    const dcaValue   = dcaShares * current;
    const dcaReturn  = ((dcaValue - dcaCost) / dcaCost) * 100;
    const dcaAvgCost = dcaCost / dcaShares;
    const dcaCAGR    = years > 0
      ? (Math.pow(dcaValue / dcaCost, 1 / years) - 1) * 100 : 0;

    // ── Winner ──
    const lsWins  = lsReturn >= dcaReturn;
    const alpha   = lsReturn - dcaReturn;

    // ── Real Portfolio ──
    const isTH   = ticker.endsWith(".BK");
    const cur    = isTH ? "฿" : "$";
    const holdings = isTH
      ? getHoldings(SHEETS.TH_HOLD)
      : getHoldings(SHEETS.US_HOLD);
    const actual = holdings.find(
      h => h.ticker.toUpperCase() === ticker.replace(".BK", "")
    );

    // ── Benchmark (QQQ สำหรับ US, SET สำหรับ TH) ──
    const benchTicker  = isTH ? "^SET.BK" : "QQQ";
    const benchHistory = fetchYahooHistory(benchTicker, 365);
    let benchReturn = 0;
    if (benchHistory && benchHistory.length >= 2) {
      const bf = benchHistory[0];
      const bl = benchHistory[benchHistory.length - 1];
      benchReturn = ((bl.close - bf.close) / bf.close) * 100;
    }
    const benchName    = isTH ? "SET Index" : "QQQ";
    const benchAlpha   = lsReturn - benchReturn;

    // ── Dates ──
    const fromStr = Utilities.formatDate(
      first.date, "Asia/Bangkok", "dd/MM/yyyy"
    );
    const toStr = Utilities.formatDate(
      last.date, "Asia/Bangkok", "dd/MM/yyyy"
    );

    // ── Insight ──
    let insight = "";
    if (lsReturn > dcaReturn * 1.5) {
      insight = ticker + " เป็นหุ้น Momentum สูง\n" +
        "Lump Sum ให้ผลตอบแทนเหนือ DCA อย่างมาก";
    } else if (dcaReturn > lsReturn) {
      insight = ticker + " เหมาะกับ DCA\n" +
        "ราคาผันผวนสูง การเฉลี่ยต้นทุนช่วยลดความเสี่ยง";
    } else {
      insight = ticker + " ทั้งสองกลยุทธ์ให้ผลใกล้เคียงกัน";
    }

    // ── Build Message ──
    let msg =
      "📊 BACKTEST — " + ticker + "\n" +
      "🕐 " + Utilities.formatDate(
        new Date(), "Asia/Bangkok", "dd/MM/yyyy"
      ) + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "📅 Period\n" +
      fromStr + " → " + toStr + "\n" +
      Math.round(days) + " Days\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "💰 Strategy #1 : Lump Sum\n\n" +
      "Buy Price      : " + cur + fmt(buyAt)   + "\n" +
      "Current Price  : " + cur + fmt(current) + "\n\n" +
      plEmoji(lsReturn) +
        " Return      : " + signPct(lsReturn)  + "\n" +
      "🚀 CAGR        : " + signPct(lsCAGR)   + "/Year\n\n" +
      "📈 Highest     : " + cur + fmt(maxClose) + "\n" +
      "📉 Lowest      : " + cur + fmt(minClose) + "\n" +
      "📉 Max Drawdown: " + fmt(maxDD) + "%\n\n" +    // ✅ ถูกแล้ว

      "━━━━━━━━━━━━\n\n" +
      "📦 Strategy #2 : Weekly DCA\n" +
      "(" + cur + fmt(weeklyBudget) + "/สัปดาห์)\n\n" +
      "Invested       : " + cur + fmt(dcaCost)    + "\n" +
      "Portfolio Value: " + cur + fmt(dcaValue)   + "\n\n" +
      "Avg Cost       : " + cur + fmt(dcaAvgCost) + "\n" +
      "Shares         : " + fmt(dcaShares)        + "\n\n" +
      plEmoji(dcaReturn) +
        " Return DCA  : " + signPct(dcaReturn) + "\n" +
      "🚀 CAGR DCA    : " + signPct(dcaCAGR) + "/Year\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "🏆 Strategy Comparison\n\n" +
      "🥇 Winner      : " + (lsWins ? "Lump Sum" : "DCA") + "\n" +
      "Lump Sum       : " + signPct(lsReturn)  + "\n" +
      "DCA            : " + signPct(dcaReturn) + "\n" +
      "Alpha          : " + signPct(Math.abs(alpha)) + "\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "📈 Benchmark (" + benchName + ")\n\n" +
      benchName + " Return : " + signPct(benchReturn)  + "\n" +
      ticker    + " Return : " + signPct(lsReturn)     + "\n" +
      (benchAlpha >= 0 ? "🏆 Outperform" : "📉 Underperform") +
        " : " + signPct(Math.abs(benchAlpha)) + "\n\n";

    // Real Portfolio (ถ้ามีในพอร์ต)
    if (actual && actual.sharesRemain > 0) {
      const actualPLPct = actual.avgCost > 0
        ? ((actual.priceNow - actual.avgCost) / actual.avgCost) * 100 : 0;
      const actualProfit = (actual.priceNow - actual.avgCost) * actual.sharesRemain;

      msg +=
        "━━━━━━━━━━━━\n\n" +
        "📋 Real Portfolio\n\n" +
        "Avg Cost       : " + cur + fmt(actual.avgCost)    + "\n" +
        "Current Price  : " + cur + fmt(actual.priceNow)   + "\n\n" +
        "Shares         : " + fmt(actual.sharesRemain)      + "\n" +
        "Portfolio Value: " + cur + fmt(actual.valueNow)   + "\n\n" +
        plEmoji(actualPLPct) +
          " P&L         : " + signPct(actualPLPct) + "\n" +
        "💰 Profit      : " + signStr(cur, actualProfit)   + "\n\n";
    }

    msg +=
      "━━━━━━━━━━━━\n\n" +
      "💡 Insight\n" + insight;

    sendTelegramSafe(msg, "BACKTEST", ticker);

  } catch (e) {
    sendTelegramError("cmdBacktest", e);
  }
}




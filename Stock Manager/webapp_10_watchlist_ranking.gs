/**
 * ============================================================
 * webapp_10_watchlist_ranking.gs
 * ------------------------------------------------------------
 * ประกอบข้อมูลสำหรับหน้า "Watchlist Analysis" (Decision View)
 * เรียกจาก client ด้วย: google.script.run.getWatchlistRanking()
 *
 * รอบนี้อ่าน webapp_07_watchlist.gs จริงแล้ว — เปลี่ยนมา reuse
 * getWatchlistData() + getWatchlistAnalysis() ตรงๆ แทนการเดา schema
 * ชีตเองและ fetch Yahoo ซ้ำ (ตัดโค้ดที่เดาไว้รอบก่อนทิ้งทั้งหมด)
 *
 * แหล่งข้อมูลที่ใช้จริงตอนนี้:
 *  - getWatchlistData()       → รายการ + market + targetPrice(เป้าซื้อ) +
 *                                supportPrice + zone + diffFromTargetPct
 *  - getFastSignal(ticker)    → ถ้าหุ้นถืออยู่จริงใน ATR Portfolio
 *  - getWatchlistAnalysis()   → ถ้ายังไม่ถือ (MA20/50/200+RSI+Volume+Decision
 *                                เทียบกับ targetPrice/supportPrice ที่ตั้งไว้)
 *
 * ⚠️ ยังไม่มีข้อมูล Fundamental (Revenue/EPS/ROE) / Valuation แบบ Fair Value
 *    (มีแต่ targetPrice ที่ผู้ใช้ตั้งเป็นราคาที่ "ตั้งใจจะซื้อ" ไม่ใช่ราคาที่คำนวณ
 *    จากปัจจัยพื้นฐาน) / Catalyst (Earnings/Analyst/News) เลย — ปล่อยเป็น
 *    'ไม่มีข้อมูล' และตัดออกจากสูตรคะแนน (ดู _wlrTotalScore) ตามที่คุยกันไว้
 *    ว่าข้ามไปก่อนจนกว่าจะมี data source จริง
 * ============================================================
 */

function getWatchlistRanking() {
  const wl = _wlrSafeCall(getWatchlistData, { success: false, items: [] });
  if (!wl.success) {
    Logger.log('getWatchlistRanking: getWatchlistData() ล้มเหลว — ' + wl.error);
    return { count: 0, statusCounts: _wlrCountByStatus([]), stocks: [] };
  }

  const modeMap = _wlrSafeCall(getStockModeMap, {});

  const stocks = wl.items.map(function (item) {
    return _wlrBuildStockObject(item, modeMap);
  }).filter(function (s) { return s !== null; });

  stocks.sort(function (a, b) { return b.totalScore - a.totalScore; });

  return {
    count: stocks.length,
    statusCounts: _wlrCountByStatus(stocks),
    stocks: stocks
  };
}

/**
 * item มาจาก getWatchlistData().items — ได้ ticker/market/targetPrice/
 * supportPrice/note/zone/diffFromTargetPct มาพร้อมแล้ว ไม่ต้องอ่านชีตเอง
 */
function _wlrBuildStockObject(item, modeMap) {
  const symbol = item.ticker;

  // ── Path A: หุ้นที่ถืออยู่จริงใน ATR Portfolio ──
  const fast = _wlrSafeCall(function () { return getFastSignal(symbol); }, null);

  let tech, reco, heldInPortfolio;

  if (fast && fast.success) {
    heldInPortfolio = true;
    tech = _wlrMapHeldTechnical(fast);
    reco = _wlrBuildRecoFromHeld(fast);
  } else {
    // ── Path B: ยังไม่ถือ — ใช้ getWatchlistAnalysis() เดิมจาก webapp_07 ──
    const wa = _wlrSafeCall(function () {
      return getWatchlistAnalysis(symbol, item.market, item.targetPrice, item.supportPrice);
    }, null);
    if (!wa || !wa.success) {
      Logger.log('_wlrBuildStockObject: ไม่มีข้อมูลวิเคราะห์สำหรับ ' + symbol + ' ข้าม (' + (wa && wa.error) + ')');
      return null; // ไม่มีข้อมูลพอ ไม่ fabricate
    }
    heldInPortfolio = false;
    tech = _wlrMapUnheldTechnical(wa);
    reco = _wlrBuildRecoFromUnheld(wa, item);
  }

  const mode = (modeMap && modeMap[symbol] && modeMap[symbol].mode) ? modeMap[symbol].mode : 'Portfolio';
  const modeLabel = mode === 'Fast' ? '⚡ Fast' : '📊 Portfolio';

  const totalScore = _wlrTotalScore(tech.technicalScore, tech.riskScore, tech.momentumScore);
  const status = _wlrStatusFromScore(totalScore);

  return {
    symbol: symbol,
    market: item.market,
    sector: 'ไม่มีข้อมูล', // ไม่มีคอลัมน์ Sector ในชีต ⭐ Watchlist จริง (มีแค่ Ticker/Market/tradeStyle)
    tradeStyle: item.tradeStyle || '',
    mode: mode,
    modeLabel: modeLabel,
    heldInPortfolio: heldInPortfolio,
    price: tech.price !== null ? tech.price : 'ไม่มีข้อมูล',
    totalScore: totalScore,
    status: status,

    technical: tech.technicalBlock,
    fundamental: { score: null, revenueGrowth: 'ไม่มีข้อมูล', epsGrowth: 'ไม่มีข้อมูล', roe: 'ไม่มีข้อมูล', debt: 'ไม่มีข้อมูล', fcf: 'ไม่มีข้อมูล', dividend: 'ไม่มีข้อมูล' },
    valuation: {
      score: null,
      current: tech.price,
      buyTrigger: item.targetPrice || 'ไม่มีข้อมูล', // ราคาที่ผู้ใช้ตั้งใจจะซื้อ — ไม่ใช่ Fair Value ที่คำนวณจากปัจจัยพื้นฐาน
      supportPrice: item.supportPrice || 'ไม่มีข้อมูล',
      diffFromTargetPct: item.diffFromTargetPct,
      zone: item.zone
    },
    risk: tech.riskBlock,
    catalyst: { earnings: 'ไม่มีข้อมูล', dividend: 'ไม่มีข้อมูล', analyst: 'ไม่มีข้อมูล', news: 'ไม่มีข้อมูล' },

    aiSummary: tech.aiSummary,
    reco: reco,
    note: item.note || ''
  };
}

/* ---------------- Path A: หุ้นที่ถืออยู่ (ใช้ getFastSignal จริง) ---------------- */
function _wlrMapHeldTechnical(fast) {
  const technicalScore = (fast.trendStars / 5) * 100;
  const riskScore = (fast.riskStars / 5) * 100;
  const momentumScore = (fast.momentumStars / 5) * 100;

  return {
    price: fast.price,
    technicalScore: technicalScore,
    riskScore: riskScore,
    momentumScore: momentumScore,
    technicalBlock: {
      score: Math.round(technicalScore),
      ema5: fast.ema5, ema20: fast.ema20,
      trend: fast.trendLabel,
      rsi: _wlrRound(fast.rsi, 1),
      volume: _wlrRound(fast.volRatio, 2),
      decision: fast.decision
    },
    riskBlock: {
      score: Math.round(riskScore),
      atr: _wlrRound(fast.atrVal, 2),
      stopLine: _wlrRound(fast.stopLine, 2),
      belowStop: fast.belowStop,
      level: fast.riskStars >= 4 ? 'Low' : fast.riskStars >= 3 ? 'Medium' : 'High'
    },
    aiSummary: [fast.decisionNote].concat(fast.reasonsFor || []).concat(fast.reasonsAgainst || [])
  };
}

function _wlrBuildRecoFromHeld(fast) {
  if (fast.belowStop || fast.decision === 'EXIT') {
    return { buyRange: '-', stopLoss: _wlrRound(fast.stopLine, 2), target1: '-', target2: '-', riskReward: '-',
             note: 'หลุด Stop แล้ว — ระบบแนะนำ EXIT ไม่ใช่จุดเข้าใหม่' };
  }
  const price = fast.price;
  const atr = fast.atrVal;
  const target1 = _wlrRound(price + atr * 2, 2);
  const target2 = _wlrRound(price + atr * 4, 2);
  const riskAmt = price - fast.stopLine;
  const rewardAmt = target2 - price;
  const rr = riskAmt > 0 ? _wlrRound(rewardAmt / riskAmt, 1) : '-';
  return {
    buyRange: _wlrRound(price * 0.99, 2) + '–' + _wlrRound(price * 1.005, 2),
    stopLoss: _wlrRound(fast.stopLine, 2),
    target1: target1, target2: target2, riskReward: rr,
    note: 'Target คำนวณจาก ATR (price + ATR×2 / ATR×4) ไม่ใช่ Fair Value เพราะยังไม่มีข้อมูล Valuation'
  };
}

/* ---------------- Path B: ยังไม่ถือ — ใช้ getWatchlistAnalysis() เดิม ---------------- */
function _wlrClassToScore(cls) {
  if (cls === 'safe') return 90;
  if (cls === 'stop') return 20;
  return 55; // 'warn' หรือไม่ทราบค่า
}

function _wlrMapUnheldTechnical(wa) {
  const technicalScore = _wlrClassToScore(wa.trendClass);
  const momentumScore = Math.round((_wlrClassToScore(wa.rsiClass) + _wlrClassToScore(wa.volClass)) / 2);

  // Risk: ใช้ supportPrice ที่ผู้ใช้ตั้งไว้เอง (ข้อมูลจริง ไม่ใช่ประมาณ) แทน ATR
  // ยิ่งห่างจาก support มาก = ยังมีที่ให้ราคาขยับก่อนหลุดแนวรับ = risk score สูง (ดี)
  let riskScore = 50, riskNote = 'ไม่มี Support Price ให้เทียบ (ยังไม่ได้ตั้งไว้ใน Watchlist)';
  if (wa.supportPrice) {
    const distPct = ((wa.price - wa.supportPrice) / wa.price) * 100;
    riskScore = distPct > 15 ? 90 : distPct > 10 ? 75 : distPct > 5 ? 60 : distPct > 0 ? 40 : 15;
    riskNote = 'ห่างจาก Support ' + _wlrRound(distPct, 1) + '%';
  }

  return {
    price: wa.price,
    technicalScore: technicalScore,
    riskScore: riskScore,
    momentumScore: momentumScore,
    technicalBlock: {
      score: Math.round(technicalScore),
      ma20: _wlrRound(wa.ma20, 2), ma50: _wlrRound(wa.ma50, 2), ma200: _wlrRound(wa.ma200, 2),
      trend: wa.trendSignal,
      rsi: _wlrRound(wa.rsi, 1),
      volume: (wa.volAvg && wa.volNow) ? _wlrRound(wa.volNow / wa.volAvg, 2) : 'ไม่มีข้อมูล',
      decision: wa.decision
    },
    riskBlock: {
      score: Math.round(riskScore),
      atr: 'ไม่มีข้อมูล (ยังไม่ถือ — ใช้ Support Price แทน)',
      stopLine: wa.supportPrice || 'ไม่มีข้อมูล',
      belowStop: null,
      level: riskScore >= 70 ? 'Low' : riskScore >= 45 ? 'Medium' : 'High',
      note: riskNote
    },
    aiSummary: [wa.trendSignal, wa.rsiSignal, wa.volSignal].concat(wa.reasons || []).concat(wa.warnings || [])
  };
}

function _wlrBuildRecoFromUnheld(wa, item) {
  if (wa.decClass === 'stop') {
    return {
      buyRange: '-', stopLoss: item.supportPrice || '-', target1: '-', target2: '-', riskReward: '-',
      note: 'ยังไม่เหมาะเข้าซื้อตามระบบวิเคราะห์ (' + wa.decision + ')'
    };
  }
  return {
    buyRange: item.targetPrice ? (_wlrRound(item.targetPrice * 0.99, 2) + '–' + _wlrRound(item.targetPrice * 1.01, 2)) : 'ไม่มีข้อมูล',
    stopLoss: item.supportPrice || 'ไม่มีข้อมูล',
    target1: 'ไม่มีข้อมูล', target2: 'ไม่มีข้อมูล', riskReward: '-',
    note: wa.decision + ' — เป้าซื้อ/Support มาจากที่ตั้งไว้ใน Watchlist เอง ยังไม่มีเป้าทำกำไร (Take-Profit) เพราะยังไม่ได้ถือ'
  };
}

/* ---------------- คะแนนรวม ---------------- */
/**
 * สูตรเดิม: Technical35 + Fundamental30 + Valuation20 + Risk10 + Momentum5
 * ตัด Fundamental/Valuation ออก (ไม่มี data source จริง — ตามที่คุยกันไว้ว่า
 * ข้ามไปก่อน) แล้ว rescale 3 หมวดที่เหลือ (35+10+5=50) ให้เต็ม 100:
 *   Technical 70% , Risk 20% , Momentum 10%
 */
function _wlrTotalScore(technicalScore, riskScore, momentumScore) {
  const total = (technicalScore * 0.70) + (riskScore * 0.20) + (momentumScore * 0.10);
  return Math.round(total);
}

function _wlrStatusFromScore(score) {
  if (score >= 75) return 'buy';
  if (score >= 40) return 'wait';
  return 'avoid';
}

/* ---------------- Helpers ---------------- */
function _wlrCountByStatus(stocks) {
  return {
    buy: stocks.filter(function (s) { return s.status === 'buy'; }).length,
    wait: stocks.filter(function (s) { return s.status === 'wait'; }).length,
    avoid: stocks.filter(function (s) { return s.status === 'avoid'; }).length,
    breakout: 0 // getFastSignal()/getWatchlistAnalysis() ไม่มีสัญญาณ Breakout แยกไว้
  };
}

function _wlrRound(v, digits) {
  if (v === null || v === undefined || isNaN(v)) return null;
  const f = Math.pow(10, digits || 0);
  return Math.round(v * f) / f;
}

function _wlrSafeCall(fn, fallback) {
  try { return fn(); } catch (e) { Logger.log('_wlrSafeCall error: ' + e.message); return fallback; }
}

/** รันทดสอบด้วยตนเองใน Apps Script Editor แล้วดู Logger */
function testWatchlistRanking() {
  Logger.log(JSON.stringify(getWatchlistRanking(), null, 2));
}

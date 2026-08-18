// ============================================================
// webapp_03c_stock_detail.gs — หน้าวิเคราะห์หุ้นรายตัว (ละเอียด)
// แบ่งเป็น 2 สายตาม StockMode: ⚡ Fast Analysis / 📊 Portfolio Analysis
// เรียกจากปุ่ม "วิเคราะห์" ในการ์ดของหน้า ภาพรวมสัญญาณ และ Rebalance ภาพรวม
//
// ⚠️ ไม่คำนวณ EMA/RSI/ATR ซ้ำเอง — ดึงจาก getFastSignal()/getHoldingsData()/
//    getRebalanceOverviewData() ที่มีอยู่แล้วเสมอ (single source of truth)
// ⚠️ ต้องมี: webapp_00b_helpers.gs, webapp_05_settings.gs (getEffectiveRiskParams),
//           webapp_02_holdings.gs, webapp_08_rebalance.gs, webapp_03_analyze.gs (getFastSignal)
// ============================================================

// ══════════════════════════════════════════════════════════
// MAIN ENTRY — เรียกจากหน้าเว็บ ครั้งเดียวได้ข้อมูลครบตามสายของหุ้นนั้น
// ══════════════════════════════════════════════════════════
function getStockDetailData(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const modeMap = getStockModeMap();
    const cfg = modeMap[ticker];
    if (!cfg) return { success: false, error: 'ไม่พบ ' + ticker + ' ในชีต StockMode — เพิ่มข้อมูลก่อนวิเคราะห์' };

    const isFast = cfg.mode === 'Fast';
    return isFast
      ? _buildFastDetailData(ticker, market, cfg)
      : _buildPortfolioDetailData(ticker, market, cfg);
  } catch (e) {
    logError('getStockDetailData', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// ⚡ FAST ANALYSIS — แปลง getFastSignal() ให้ตรงโครง spec
// ══════════════════════════════════════════════════════════
function _buildFastDetailData(ticker, market, cfg) {
  const sig = getFastSignal(ticker);
  if (!sig || !sig.success) {
    return { success: false, error: (sig && sig.error) || 'วิเคราะห์ไม่สำเร็จ' };
  }

  const cur = sig.cur;
  const action = _mapDecisionToSignal(sig.decisionClass); // reuse ตัวเดียวกับหน้าภาพรวมสัญญาณ — กันฉลากไม่ตรงกันข้ามหน้า



  // ── Take Profit price: ไม่มีใน getFastSignal ตรงๆ คำนวณเองจาก eff.takeProfitPct (ไม่แก้ไฟล์เดิม) ──
  const eff = getEffectiveRiskParams(ticker);
  const takeProfitPrice = sig.buyPrice > 0 ? sig.buyPrice * (1 + eff.takeProfitPct) : null;

  // ── Risk/Reward: (กำไรเป้าหมาย) / (ความเสี่ยงถ้าหลุด stop) — null ถ้าราคาต่ำกว่า/เท่า stop อยู่แล้ว ──
  const riskAmount = sig.price - sig.stopLine;
  const rewardAmount = takeProfitPrice ? takeProfitPrice - sig.price : null;
  const riskReward = (riskAmount > 0 && rewardAmount !== null) ? rewardAmount / riskAmount : null;
  const distancePct = sig.price > 0 ? ((sig.price - sig.stopLine) / sig.price) * 100 : null;

  // ── Momentum/Risk แบบสั้น จาก stars ที่มีอยู่แล้ว (ไม่คำนวณซ้ำ) ──
  const momentumLabel = sig.momentumStars >= 4 ? 'Strong' : (sig.momentumStars === 3 ? 'Neutral' : 'Weak');
  const momentumClass = sig.momentumStars >= 4 ? 'safe' : (sig.momentumStars === 3 ? 'warn' : 'stop');
  const riskLabel = sig.riskStars >= 4 ? 'Low' : (sig.riskStars === 3 ? 'Medium' : 'High');
  const riskClass = sig.riskStars >= 4 ? 'safe' : (sig.riskStars === 3 ? 'warn' : 'stop');

  const realized = _getRealizedPnLForTicker(ticker, sig.isTH ? 'TH' : 'US'); // ← เพิ่มบรรทัดนี้ ก่อน return

  // ── Technical Checklist — พื้นฐานที่มีอยู่แล้วเสมอ (EMA5/20, RSI, ราคา>EMA20) ──
  const checklist = [
    { ok: sig.ema5 > sig.ema20, label: 'EMA5 > EMA20 (แนวโน้มระยะสั้น)' },
    { ok: sig.rsi >= 50 && sig.rsi <= 70, label: 'RSI ' + sig.rsi.toFixed(1) + ' อยู่ในโซนแข็งแรง (50-70)', warn: sig.rsi > 70 },
    { ok: sig.price >= sig.ema20, label: 'ราคา > EMA20' }
  ];


  // ── ต่อ EMA50/MACD จาก Daily_Close_Log (ไฟล์ภายนอก) ถ้ามีข้อมูลพอ — ไม่มีก็ไม่ใส่ ไม่เดา ──
  let checklistCaption = 'ระบบ Fast Trade ปัจจุบันคำนวณจาก EMA5/EMA20 + RSI + Volume — ยังไม่รวม MACD และ EMA50';
  try {
    const ext = getExtendedTechnicals(ticker);
    if (ext.available) {
      if (ext.ema50 !== undefined) {
        checklist.push({ ok: ext.ema20AboveEMA50, label: 'EMA20 > EMA50 (แนวโน้มระยะกลาง)' });
        checklist.push({ ok: ext.priceAboveEMA50, label: 'ราคา > EMA50' });
      }
      if (ext.macdLine !== undefined) {
        checklist.push({ ok: ext.macdBullish, label: 'MACD Bullish (MACD Line > Signal Line)' });
      }
      checklistCaption = `เพิ่มเติมจาก Daily_Close_Log (${ext.daysAvailable} วันย้อนหลัง)` +
        (ext.ema50Note ? ` — EMA50 ${ext.ema50Note}` : '') +
        (ext.macdNote ? ` — MACD ${ext.macdNote}` : '');
    } else {
      checklistCaption += ' · ' + ext.reason;
    }
  } catch (e) {
    // ไม่มีไฟล์ webapp_09_external_history.gs หรือยังไม่ได้ตั้งค่า — ใช้ checklist พื้นฐานต่อไปได้ปกติ ไม่ให้ทั้งหน้าพัง
  }

  // ── หาแนวรับ (Swing Low) จาก Daily_Close_Log — เป็นข้อมูลอ้างอิงเสริม ไม่แทนที่ Hard Stop เดิม
  //    (ตามหลักการ "Swing Low → Stop Loss" ของสาย Fast) ──
  const support = _getSupportSafe(ticker);
  // ── หาแนวต้าน (Swing High + All-Time High) — เป็นข้อมูลอ้างอิงเสริมเทียบกับ Take Profit ──
  const resistance = _getResistanceSafe(ticker);

  return {
    success: true, lane: 'fast', ticker, market: sig.isTH ? 'TH' : 'US', cur,
   
    hero: {
      price: sig.price, plPct: sig.plPct, action,
      actionLabel: { buy: 'BUY', watch: 'WATCH', sell: 'SELL' }[action],
      strategy: '⚡ Fast Trade',
      rsi: sig.rsi,
      trend: { label: sig.trendLabel, cls: sig.trendClass }
    },


    tradePlan: {
      entry: sig.buyPrice,       // ราคาที่เข้าซื้อจริง (ถืออยู่แล้ว ไม่ใช่ราคาที่วางแผนจะเข้าใหม่)
      entryIsActual: true,       // ใช้บอก frontend ว่าเลขนี้คือ "ซื้อแล้ว" ไม่ใช่ "แผนในอนาคต"
      takeProfit: takeProfitPrice,
      trailingStop: sig.trailingStopPrice,
      hardStop: sig.cutStopPrice,
      shares: sig.shares

    },
    signalSummary: {
      trend: { label: sig.trendLabel, cls: sig.trendClass },
      momentum: { label: momentumLabel, cls: momentumClass },
      volume: { label: sig.emaConditions[2].text, cls: sig.emaConditions[2].cls },
      risk: { label: riskLabel, cls: riskClass }
    },
    checklist,
    checklistCaption,
    support, // { available, currentPrice, rolling20Low, rolling50Low, swings:[{date,low}], nearestSupport (number), distancePct } หรือ { available:false, reason }
    resistance, // { available, currentPrice, rolling20High, rolling50High, swings:[{date,high}], allTimeHigh:{price,date}, nearestResistance (number), distancePct }
    riskManagement: {
      currentPrice: sig.price, trailingStop: sig.trailingStopPrice,
      distancePct, riskReward
    },
    recommendation: {
      verdict: sig.decision, verdictClass: sig.decisionClass, note: sig.decisionNote,
      reasons: sig.reasonsFor, warnings: sig.reasonsAgainst,
      actions: sig.actionPlan ? [sig.actionPlan.action] : _defaultFastActions(sig.decision)
    },
  
  feeCard: _attachFeeInsights(
  getStockFeeCard(ticker, market),
  sig.buyPrice, sig.shares,
  (sig.price - sig.buyPrice) * sig.shares,   // unrealized
  realized.found ? realized.netPL : 0        // realized
),
updatedAt: sig.updatedAt
  };
}

function _defaultFastActions(decision) {
  if (decision === 'EXIT') return ['ขายทั้งหมด'];
  if (decision === 'HOLD') return ['ถือต่อ', 'ยังไม่ซื้อเพิ่ม'];
  return ['ถือต่อ', 'เฝ้าดูสัญญาณใกล้ชิด'];
}

// ══════════════════════════════════════════════════════════
// 📊 PORTFOLIO ANALYSIS — รวมจาก Holdings + Rebalance + ปันผลรายตัว
// ══════════════════════════════════════════════════════════
function _buildPortfolioDetailData(ticker, market, cfg) {
  const h = getHoldingsData();
  const arr = (market === 'TH') ? h.th : h.us;
  const holding = (arr || []).find(x => x.ticker === ticker);
  if (!holding) return { success: false, error: 'ไม่พบ ' + ticker + ' ในรายการถือครองปัจจุบัน (Holdings)' };

  const cur = market === 'TH' ? '฿' : '$';

    // ── ดึงราคาสดจริงผ่าน lookupTickerPrice() แทนอ่าน holding.priceNow จากเซลล์ในชีต Holdings ──
  // (เซลล์นั้นอัปเดตแค่ตอนกดปุ่ม "อัปเดตราคา" เท่านั้น ทำให้ปุ่ม ▶️ auto-refresh ทุก 30 วิ เห็นราคาเดิมซ้ำ)
  // แล้วเขียนราคาใหม่กลับเข้าชีตจริงด้วย ให้หน้า Holdings/Summary/Rebalance เห็นตรงกัน ไม่ใช่แค่หน้านี้
  try {
    const live = lookupTickerPrice(ticker, market);
    if (live && live.success && live.price > 0) {
      const qty = holding.sharesRemain;
      holding.priceNow = live.price;
      holding.valueNow = live.price * qty;
      holding.unrealizedPL = holding.valueNow - holding.totalCost;
      holding.unrealizedPct = holding.totalCost > 0 ? (holding.unrealizedPL / holding.totalCost) : 0;

      const sheetName = (market === 'TH') ? SHEETS.TH_HOLD : SHEETS.US_HOLD;
      writeLivePriceToHoldingsSheet(sheetName, ticker, live.price); // ← เขียนกลับเข้าชีตจริง
    }
  } catch (e) {
    // lookupTickerPrice error — ใช้ราคาจากชีตเดิมต่อไปได้ปกติ ไม่ให้ทั้งหน้าพัง
  }


  // ── ดึงสัดส่วนพอร์ต (current/target weight) จาก getRebalanceOverviewData() — ไม่คำนวณ THB conversion ซ้ำ ──
  const rebal = getRebalanceOverviewData();
  let weightInfo = null;
  if (rebal.success) weightInfo = (rebal.stocks || []).find(x => x.ticker === ticker && x.market === market);

  // ── วันที่ถือครอง (จากไม้แรกที่ยังไม่ขายหมด) + ปันผลสะสมของหุ้นนี้ + Realized P&L สุทธิ ──
  const firstBuyDate = _getFirstBuyDateForTicker(ticker, market);
  const holdingDays = firstBuyDate ? Math.floor((new Date() - firstBuyDate) / (1000 * 60 * 60 * 24)) : null;
  const dividendReceived = _getDividendSumForTicker(ticker, market);
  const realized = _getRealizedPnLForTicker(ticker, market);

  // ── Rebalance action: ต้องซื้อ/ขายเท่าไหร่เพื่อกลับเข้าเป้าหมาย (คำนวณจากมูลค่ารวมพอร์ตจริง) ──
  let rebalanceAction = null;
  if (weightInfo && rebal.totalValueTHB > 0) {
    const targetValueTHB = rebal.totalValueTHB * (weightInfo.targetWeightPct / 100);
    const diffTHB = targetValueTHB - weightInfo.valueTHB;
    rebalanceAction = {
      currentWeightPct: weightInfo.currentWeightPct, targetWeightPct: weightInfo.targetWeightPct,
      diffTHB, needBuy: diffTHB > 0, amountTHB: Math.abs(diffTHB)
    };
  }

  const decision = weightInfo
    ? (weightInfo.overConcentration ? 'REDUCE' : (weightInfo.needsRebalance ? (weightInfo.weightDiffPct < 0 ? 'BUY_MORE' : 'REDUCE') : 'HOLD'))
    : 'HOLD';
  const decisionClass = decision === 'REDUCE' ? 'stop' : (decision === 'BUY_MORE' ? 'safe' : 'warn');

  return {
    success: true, lane: 'portfolio', ticker, market, cur,
    hero: {
      decision, decisionClass,
      currentWeightPct: weightInfo ? weightInfo.currentWeightPct : null,
      targetWeightPct: weightInfo ? weightInfo.targetWeightPct : null,
      isEstimatedWeight: weightInfo ? weightInfo.isEstimated : null
    },

   

    positionSummary: {
      avgCost: holding.avgCost, currentPrice: holding.priceNow,
      returnPct: holding.unrealizedPct * 100, unrealizedPL: holding.unrealizedPL
    },
    allocation: weightInfo ? {
      currentPct: weightInfo.currentWeightPct, targetPct: weightInfo.targetWeightPct,
      diffPct: weightInfo.weightDiffPct, isEstimated: weightInfo.isEstimated
    } : null,
    holdingInfo: {
      holdingDays, shares: holding.sharesRemain, currentValue: holding.valueNow,
      cost: holding.totalCost, market
    },
    performance: {
      capitalGain: holding.unrealizedPL,
      dividend: dividendReceived,
      realizedProfit: realized.found ? realized.netPL : null,
      realizedProfitNote: realized.found
        ? `รวมจาก ${realized.numRows} รายการขาย (${realized.numSales} ครั้ง) — ยอดสุทธิหลังหักค่าธรรมเนียมแล้ว`
        : 'ยังไม่เคยขายหุ้นตัวนี้เลย — ไม่มี Realized P&L'
    },
    rebalance: rebalanceAction,
    support: _getSupportSafe(ticker), // Swing Low เดียวกับสาย Fast — แต่ตีความเป็น "โซนรอซื้อเพิ่ม" แทน Stop Loss
    resistance: _getResistanceSafe(ticker), // Swing High/ATH — ตีความเป็น "โซนอาจพิจารณาขายทำกำไรบางส่วน"
    recommendation: _buildPortfolioRecommendation(decision, weightInfo, holding),
feeCard: _attachFeeInsights(
  getStockFeeCard(ticker, market),
  holding.avgCost, holding.sharesRemain,
  holding.unrealizedPL,                       // unrealized
  realized.found ? realized.netPL : 0         // realized
),
updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')

  };
}

// ── เรียก getSupportLevels() แบบกันพัง เผื่อไฟล์ webapp_09_external_history.gs ไม่มี/ยังไม่ตั้งค่า ──
function _getSupportSafe(ticker) {
  try {
    return getSupportLevels(ticker);
  } catch (e) {
    return { available: false, reason: 'ไม่มีข้อมูล (เช็คว่ามีไฟล์ webapp_09_external_history.gs และตั้งค่า EXTERNAL_LOG_SHEET_ID แล้วหรือยัง)' };
  }
}

// ── เรียก getResistanceLevels() แบบกันพัง เช่นเดียวกับ _getSupportSafe ──
function _getResistanceSafe(ticker) {
  try {
    return getResistanceLevels(ticker);
  } catch (e) {
    return { available: false, reason: 'ไม่มีข้อมูล (เช็คว่ามีไฟล์ webapp_09_external_history.gs และตั้งค่า EXTERNAL_LOG_SHEET_ID แล้วหรือยัง)' };
  }
}

function _buildPortfolioRecommendation(decision, weightInfo, holding) {
  const reasons = [];
  const profitPct = holding.unrealizedPct * 100;
  reasons.push((profitPct >= 0 ? 'กำไร ' : 'ขาดทุน ') + Math.abs(profitPct).toFixed(1) + '%');

  if (weightInfo) {
    if (weightInfo.overConcentration) reasons.push('สัดส่วนเกิน Concentration Limit — เสี่ยงกระจุกตัวสูง');
    else if (weightInfo.needsRebalance) reasons.push(weightInfo.weightDiffPct < 0 ? 'สัดส่วนต่ำกว่าเป้าหมาย' : 'สัดส่วนเกินเป้าหมาย');
    else reasons.push('สัดส่วนใกล้เคียงเป้าหมาย');
  }

  const actions = decision === 'REDUCE' ? ['พิจารณาลดสัดส่วน']
    : decision === 'BUY_MORE' ? ['ถือต่อ', 'ซื้อเพิ่มเมื่อย่อ']
    : ['ถือต่อ'];

  return { verdict: decision, reasons, actions };
}

// ── หาผลรวมปันผลของหุ้นตัวนี้ จากชีตบันทึกปันผล (SHEETS.DIV) — กรองด้วย ticker ตรงๆ ──
function _getDividendSumForTicker(ticker, market) {
  try {
    const sheet = getSheet(SHEETS.DIV);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.DIV) return 0;

    const numRows = lastRow - START_ROW.DIV + 1;
    const rows = sheet.getRange(START_ROW.DIV, 1, numRows, 17).getValues();

    let sum = 0;
    rows.forEach(row => {
      const rTicker = String(row[DIV_COL.TICKER - 1] || '').trim().toUpperCase();
      if (rTicker === ticker) sum += Number(row[DIV_COL.NET_THB - 1]) || 0;
    });
    return sum;
  } catch (e) {
    logError('_getDividendSumForTicker', e);
    return 0;
  }
}

// ── หาวันที่ซื้อไม้แรกของหุ้นตัวนี้ (ที่ยังไม่ขายหมด) จาก Transaction Log ──
function _getFirstBuyDateForTicker(ticker, market) {
  try {
    const logSheetName = (market === 'TH') ? SHEETS.TH_TRANS : SHEETS.US_TRANS;
    const rows = getSheet(logSheetName).getDataRange().getValues();
    const buys = rows
      .filter(r => String(r[2] || '').trim().toUpperCase().replace(/_C\d+$/i, '') === ticker && r[3] === 'ซื้อ' && r[1] instanceof Date)
      .map(r => r[1])
      .sort((a, b) => a - b);
    return buys.length ? buys[0] : null;
  } catch (e) {
    logError('_getFirstBuyDateForTicker', e);
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// Realized P&L รายตัว — อ่านจากตารางจริงในชีต SHEETS.US_REAL / SHEETS.TH_REAL
// (ยืนยันโครงสร้างจากภาพจริง: header แถว 6, ข้อมูลเริ่มแถว 7, คอลัมน์ B-L)
// รวมทุกแถวที่เป็นหุ้นตัวเดียวกัน (ตัด suffix _c1/_c2/_c3 ทิ้งก่อนเทียบ — เช่น
// NVDA, NVDA_c1, NVDA_c2 คือหุ้นตัวเดียวกันแค่คนละไม้/คนละรอบขาย) ──
const REALIZED_TABLE_START_ROW = 7;
const REALIZED_TABLE_COL = {
  TICKER: 1,       // B (index 0 หลัง getRange เริ่มจาก col B)
  AVG_COST: 2,      // C
  SALE_VALUE: 3,    // D
  COST_SOLD: 4,     // E
  REALIZED_PL: 5,   // F (ก่อนหักค่าธรรมเนียม)
  PL_PCT: 6,        // G
  NUM_SALES: 7,     // H
  STATUS: 8,        // I
  FEES: 9,          // J
  NET_PL: 10,       // K ← ใช้ตัวนี้เป็นค่าจริง (สุทธิหลังหักค่าธรรมเนียมแล้ว)
  NET_PL_PCT: 11    // L
};

function _getRealizedPnLForTicker(ticker, market) {
  try {
    const sheetName = (market === 'TH') ? SHEETS.TH_REAL : SHEETS.US_REAL;
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < REALIZED_TABLE_START_ROW) return { found: false };

    const numRows = lastRow - REALIZED_TABLE_START_ROW + 1;
    // เริ่มอ่านจากคอลัมน์ B (2) ไปอีก 11 คอลัมน์ถึง L
    const rows = sheet.getRange(REALIZED_TABLE_START_ROW, 2, numRows, 11).getValues();

    let netPL = 0, numSales = 0, matchedRows = 0;
    rows.forEach(row => {
      const rawTicker = String(row[REALIZED_TABLE_COL.TICKER] || '').trim().toUpperCase();
      const cleanTicker = rawTicker.replace(/_C\d+$/i, ''); // ตัด _c1, _c2, _c3... ทิ้ง
      if (cleanTicker !== ticker) return;

      const rowNetPL = Number(row[REALIZED_TABLE_COL.NET_PL]) || 0;
      const rowNumSales = Number(row[REALIZED_TABLE_COL.NUM_SALES]) || 0;
      if (rowNumSales === 0) return; // แถวที่ยังไม่เคยขาย ("—") ไม่นับ

      netPL += rowNetPL;
      numSales += rowNumSales;
      matchedRows++;
    });

    if (matchedRows === 0) return { found: false };
    return { found: true, netPL, numSales, numRows: matchedRows };
  } catch (e) {
    logError('_getRealizedPnLForTicker', e);
    return { found: false };
  }
}

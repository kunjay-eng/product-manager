// ============================================================
// webapp_08_rebalance.gs — หน้า Rebalance ภาพรวม (3 sub-tab)
// 1) Portfolio Trend  2) Target Weight  3) แผนเข้าซื้อ (Buy Plan)
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
//
// ⚠️ ต้องมี: webapp_00b_helpers.gs (getSheet, logError)
//           webapp_05_settings.gs (getStockModeMap, getRiskSettings)
//           webapp_02_holdings.gs (getHoldingsData)
//           webapp_00_main.gs (_getNextEmptyRow)
//
// ต้องรัน setupRebalanceExtras() ครั้งเดียวก่อนใช้งาน (เพิ่ม cell ใน Settings
// + สร้างชีต BuyPlan) — เลือกฟังก์ชันนี้จากดรอปดาวน์ Apps Script แล้ว ▶ Run
// ============================================================

// ── เพิ่ม cell ใหม่ใน Settings สำหรับ Rebalance (ต่อจาก E4:E10 เดิม) ──
const REBAL_SETTINGS = {
  FREQ_MONTHS:        'E11', // ความถี่ rebalance ตามรอบเวลา (เดือน)
  LAST_REBAL_DATE:    'E12', // วันที่ rebalance ล่าสุด
  DEVIATION_PCT:       'E13'  // % เบี่ยงจาก target ที่ถือว่า "ต้อง rebalance"
};

const BUY_PLAN_SHEET = {
  NAME: 'BuyPlan',
  START_ROW: 7
  // คอลัมน์ (1-based): A=Ticker B=Market C=PlanType(price/time) D=Budget รวม
  // E=Leg1% F=Leg1TriggerPct(ห่างจาก entry แรก) G=Leg2% H=Leg2TriggerPct
  // I=Leg3% J=Leg3TriggerPct K=DCA ความถี่(วัน) L=DCA จำนวนเงิน/ครั้ง
  // M=วันที่เริ่มแผน N=หมายเหตุ
};

// ══════════════════════════════════════════════════════════
// ติดตั้งครั้งเดียว — เพิ่ม cell Rebalance ใน Settings + สร้างชีต BuyPlan
// ══════════════════════════════════════════════════════════
function setupRebalanceExtras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = '';

  // ── เพิ่ม cell ใน Settings ถ้ายังไม่มี (เช็คจาก label ว่าง) ──
  try {
    const sheet = getSheet(SETTINGS_SHEET.SHEET);
    if (!sheet.getRange('B11').getValue()) {
      sheet.getRange('B11').setValue('ความถี่ Rebalance ตามรอบเวลา (เดือน)');
      sheet.getRange(REBAL_SETTINGS.FREQ_MONTHS).setValue(3);
      sheet.getRange('B12').setValue('วันที่ Rebalance ล่าสุด');
      sheet.getRange(REBAL_SETTINGS.LAST_REBAL_DATE).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
      sheet.getRange('B13').setValue('% เบี่ยงจาก Target ที่ถือว่าต้อง Rebalance');
      sheet.getRange(REBAL_SETTINGS.DEVIATION_PCT).setValue(0.05).setNumberFormat('0.00%');
      sheet.getRange('E11:E13').setBackground('#fff2cc');
      msg += '✅ เพิ่มค่า Rebalance ใน Settings แล้ว\n';
    } else {
      msg += 'ℹ️ Settings มีค่า Rebalance อยู่แล้ว ไม่ได้แก้ไข\n';
    }
  } catch (e) {
    msg += '❌ เพิ่มค่าใน Settings ไม่สำเร็จ: ' + e.message + '\n';
  }

  // ── สร้างชีต BuyPlan ถ้ายังไม่มี ──
  if (ss.getSheetByName(BUY_PLAN_SHEET.NAME)) {
    msg += 'ℹ️ ชีต BuyPlan มีอยู่แล้ว ไม่ได้แก้ไข';
  } else {
    const sheet = ss.insertSheet(BUY_PLAN_SHEET.NAME);
    sheet.getRange('B2').setValue('🛒 BuyPlan — แผนเข้าซื้อรายตัว (Price-based / Time-based DCA)')
      .setFontWeight('bold').setFontSize(12);
    sheet.getRange('B3').setValue('💡 กรอกอย่างน้อย Ticker/Market/PlanType — ช่องที่ไม่ใช้เว้นว่างได้ (เช่น price-based ไม่ต้องกรอก DCA)');

    const headers = ['Ticker', 'Market', 'PlanType (price/time)', 'งบรวม',
      'Leg1 %', 'Leg1 Trigger (ห่างจาก entry แรก %)', 'Leg2 %', 'Leg2 Trigger %',
      'Leg3 %', 'Leg3 Trigger %', 'DCA ความถี่ (วัน)', 'DCA จำนวนเงิน/ครั้ง',
      'วันที่เริ่มแผน', 'หมายเหตุ'];
    const headerRow = BUY_PLAN_SHEET.START_ROW - 1;
    sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#e8eaf6').setWrap(true);

    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['price', 'time'], true).setAllowInvalid(true).build();
    sheet.getRange(BUY_PLAN_SHEET.START_ROW, 3, 100, 1).setDataValidation(rule);
    sheet.getRange(BUY_PLAN_SHEET.START_ROW, 4, 100, 10).setBackground('#fff2cc');
    sheet.setColumnWidths(2, 13, 130);
    msg += '✅ สร้างชีต BuyPlan แล้ว';
  }

  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

// ══════════════════════════════════════════════════════════
// อ่านค่า Rebalance settings (พร้อม fallback ถ้ายังไม่ได้ setupRebalanceExtras)
// ══════════════════════════════════════════════════════════
function getRebalanceSettings() {
  const fallback = { freqMonths: 3, lastRebalanceDate: null, deviationPct: 0.05 };
  try {
    const sheet = getSheet(SETTINGS_SHEET.SHEET);
    const freq = Number(sheet.getRange(REBAL_SETTINGS.FREQ_MONTHS).getValue());
    const lastDateVal = sheet.getRange(REBAL_SETTINGS.LAST_REBAL_DATE).getValue();
    const dev = Number(sheet.getRange(REBAL_SETTINGS.DEVIATION_PCT).getValue());

    return {
      freqMonths: freq || fallback.freqMonths,
      lastRebalanceDate: (lastDateVal instanceof Date) ? lastDateVal : fallback.lastRebalanceDate,
      deviationPct: dev || fallback.deviationPct
    };
  } catch (e) {
    logError('getRebalanceSettings', e);
    return fallback;
  }
}

// ══════════════════════════════════════════════════════════
// MAIN ENTRY — เรียกจากหน้าเว็บครั้งเดียว ได้ข้อมูลครบทั้ง 3 sub-tab
// (sub-tab "แผนเข้าซื้อ" ได้แค่ "รายชื่อหุ้นให้เลือก" จากตรงนี้ ส่วนรายละเอียด
//  แผนของแต่ละตัวเรียกแยกทีหลังผ่าน getBuyPlanForTicker() ตอนผู้ใช้เลือกหุ้น)
// ══════════════════════════════════════════════════════════
function getRebalanceOverviewData() {
  try {
    const modeMap = getStockModeMap();
    const risk = getRiskSettings();
    const rebalSettings = getRebalanceSettings();

    // ── รวม holdings ทั้งหมด (US/TH) พร้อมมูลค่าเป็น THB ──
    // ⚠️ ถ้าดึงอัตราแลกเปลี่ยนไม่ได้ จุดนี้จะ throw error ทันที (ดู _getUsdThbRateOrThrow)
    //    ตั้งใจให้ทั้งฟังก์ชัน fail ชัดเจน ดีกว่าคำนวณสัดส่วน THB ด้วยเลขที่เดาขึ้นมา
    const holdingsRaw = getHoldingsData();
    const combinedResult = _combineHoldingsWithMode(holdingsRaw, modeMap);
    const combined = combinedResult.rows;

    const totalValueTHB = combined.reduce((s, x) => s + x.valueTHB, 0);
    combined.forEach(x => { x.currentWeightPct = totalValueTHB > 0 ? (x.valueTHB / totalValueTHB) * 100 : 0; });

    // ── Target Weight: ใช้ override ก่อน ไม่มีค่อย auto ไม่มีค่อยแบ่งเท่าๆ กันในกลุ่มเดียวกัน ──
    _applyTargetWeights(combined, modeMap);

    // ── สถานะ rebalance ต่อตัว ──
    const deviationPct = rebalSettings.deviationPct * 100;
    combined.forEach(x => {
      x.weightDiffPct = x.currentWeightPct - x.targetWeightPct;
      x.needsRebalance = Math.abs(x.weightDiffPct) > deviationPct;
      x.overConcentration = (x.currentWeightPct / 100) > risk.concentrationWarnPct;
    });

    // ── สรุปตามกลุ่ม (Trend Group) ──
    const groupAlloc = _summarizeByGroup(combined);

    // ── Portfolio Trend: กลุ่มไหนเอียงเกินเป้าหมายกลุ่มมากที่สุด ──
    const trend = _computePortfolioTrend(groupAlloc);

    // ── เงื่อนไขแจ้งเตือน: เวลา vs % เบี่ยง อันไหนถึงก่อน ──
    const trigger = _computeRebalanceTrigger(rebalSettings, combined);

    return {
      success: true,
      totalValueTHB,
      fxRateUsed: combinedResult.usdRate,
      totalStocks: combined.length,
      needsRebalanceCount: combined.filter(x => x.needsRebalance).length,
      overConcentrationCount: combined.filter(x => x.overConcentration).length,
      groupAllocation: groupAlloc,
      trend,
      trigger,
      stocks: combined.sort((a, b) => Math.abs(b.weightDiffPct) - Math.abs(a.weightDiffPct)), // เบี่ยงมากสุดขึ้นก่อน
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('getRebalanceOverviewData', e);
    return { success: false, error: e.message }; // e.message จะเป็นข้อความชัดเจนจาก _getUsdThbRateOrThrow ถ้าปัญหาคือ FX
  }
}

// ── ดึงอัตราแลกเปลี่ยน USD→THB จาก getFxRate() ตัวจริงใน data.gs
//    (ยืนยันชื่อฟังก์ชันแล้วจากไฟล์ data.gs จริงของโปรเจกต์ — ไม่ต้องเดาชื่ออีกต่อไป)
//    getFxRate() ภายในมี fallback 33.4 อยู่แล้วถ้าดึงจากชีต FX_Alert ไม่ได้ (ดูใน data.gs)
//    ⚠️ ยังคง throw error ที่นี่ถ้าเรียกฟังก์ชันไม่สำเร็จเลย (เช่น ชีต FX ถูกลบ/เปลี่ยนชื่อ)
//    เพื่อไม่ให้คำนวณสัดส่วน THB ต่อด้วยเลขที่ไม่รู้ที่มา
// ══════════════════════════════════════════════════════════
function _getUsdThbRateOrThrow() {
  try {
    const rate = getFxRate(); // ฟังก์ชันจริงจาก data.gs
    if (rate && !isNaN(rate) && rate > 0) return rate;
  } catch (e) {
    logError('_getUsdThbRateOrThrow', e);
  }
  throw new Error('ไม่สามารถดึงอัตราแลกเปลี่ยน USD/THB ได้ (getFxRate() ใน data.gs) — ' +
    'เช็คว่าชีต FX_Alert ยังอยู่และ FX_CELL.RATE ชี้ไปยัง cell ที่มีค่าจริง');
}

// ── รวม holdings US/TH เข้าด้วยกัน แปลงเป็น THB และผูกกับข้อมูลจาก StockMode ──
function _combineHoldingsWithMode(holdingsRaw, modeMap) {
  const result = [];
  const usdRate = _getUsdThbRateOrThrow(); // จะ throw ทันทีถ้าดึงไม่ได้ — ให้ getRebalanceOverviewData() จับแล้วแจ้ง error ตรงๆ

  (holdingsRaw.us || []).forEach(h => {
    const cfg = modeMap[h.ticker] || {};
    result.push({
      ticker: h.ticker, market: 'US',
      valueTHB: (parseFloat(h.valueNow) || 0) * usdRate,
      trendGroup: cfg.trendGroup || 'ไม่มีกลุ่ม',
      assetType: cfg.assetType || '',
      mode: cfg.mode || 'Portfolio',
      targetWeightAuto: cfg.targetWeightAuto,
      targetWeightOverride: cfg.targetWeightOverride
    });
  });
  (holdingsRaw.th || []).forEach(h => {
    const cfg = modeMap[h.ticker] || {};
    result.push({
      ticker: h.ticker, market: 'TH',
      valueTHB: parseFloat(h.valueNow) || 0,
      trendGroup: cfg.trendGroup || 'ไม่มีกลุ่ม',
      assetType: cfg.assetType || '',
      mode: cfg.mode || 'Portfolio',
      targetWeightAuto: cfg.targetWeightAuto,
      targetWeightOverride: cfg.targetWeightOverride
    });
  });
  return { rows: result, usdRate };
}

// ── กำหนด targetWeightPct ให้แต่ละตัว: override > auto(ที่บันทึกไว้) > ประมาณการแบ่งเท่ากันในกลุ่ม
//    ⚠️ กรณี fallback ("ประมาณการ") ติด isEstimated:true เสมอ — ห้าม frontend แสดงเหมือนเป็นค่าที่
//    ระบบคำนวณอัจฉริยะ เพราะจริงๆ คือการเดาคร่าวๆ (แบ่งเท่ากันในกลุ่มเดียวกัน) ไม่ใช่การวิเคราะห์จริง ──
function _applyTargetWeights(combined, modeMap) {
  // นับจำนวนตัวต่อกลุ่ม สำหรับ fallback แบ่งเท่าๆ กัน
  const groupCounts = {};
  combined.forEach(x => { groupCounts[x.trendGroup] = (groupCounts[x.trendGroup] || 0) + 1; });

  combined.forEach(x => {
    if (x.targetWeightOverride !== null && x.targetWeightOverride !== undefined) {
      x.targetWeightPct = x.targetWeightOverride;
      x.targetWeightSource = 'override';
      x.isEstimated = false;
    } else if (x.targetWeightAuto !== null && x.targetWeightAuto !== undefined) {
      x.targetWeightPct = x.targetWeightAuto;
      x.targetWeightSource = 'auto';
      x.isEstimated = false;
    } else {
      // ประมาณการ: แบ่งเท่าๆ กันภายในกลุ่มเดียวกัน โดยสมมติทุกกลุ่มมีน้ำหนักเท่ากัน (100% / จำนวนกลุ่ม / จำนวนตัวในกลุ่ม)
      // นี่คือค่าประมาณหยาบๆ ชั่วคราว ไม่ใช่ผลคำนวณจากปัจจัยพื้นฐาน/ความเสี่ยงจริง
      const numGroups = Object.keys(groupCounts).length || 1;
      x.targetWeightPct = (100 / numGroups) / groupCounts[x.trendGroup];
      x.targetWeightSource = 'fallback_equal';
      x.isEstimated = true;
    }
  });
}

// ── รวมสัดส่วนปัจจุบัน + เป้าหมาย ตามกลุ่ม (Trend Group) ──
function _summarizeByGroup(combined) {
  const groups = {};
  combined.forEach(x => {
    if (!groups[x.trendGroup]) groups[x.trendGroup] = { group: x.trendGroup, currentPct: 0, targetPct: 0 };
    groups[x.trendGroup].currentPct += x.currentWeightPct;
    groups[x.trendGroup].targetPct += x.targetWeightPct;
  });
  return Object.values(groups).sort((a, b) => b.currentPct - a.currentPct);
}

// ── สรุปแนวโน้มพอร์ต: กลุ่มที่เอียงเกินเป้าหมายมากที่สุด (บวก = เกิน, ลบ = ขาด) ──
function _computePortfolioTrend(groupAlloc) {
  if (!groupAlloc.length) return { label: 'ไม่มีข้อมูล', dominantGroup: null, diffPct: 0 };

  const sorted = [...groupAlloc].sort((a, b) => (b.currentPct - b.targetPct) - (a.currentPct - a.targetPct));
  const top = sorted[0];
  const diffPct = top.currentPct - top.targetPct;

  let label;
  if (Math.abs(diffPct) < 2) label = '⚖️ สมดุลตามแผน';
  else if (diffPct > 0) label = `📈 ${top.group}-Tilted`;
  else label = `📉 ${top.group} ต่ำกว่าเป้าหมาย`;

  return { label, dominantGroup: top.group, diffPct, currentPct: top.currentPct, targetPct: top.targetPct };
}

// ── เช็คเงื่อนไขแจ้งเตือน rebalance: เวลา vs % เบี่ยง อันไหนถึงก่อน ──
function _computeRebalanceTrigger(rebalSettings, combined) {
  const now = new Date();
  let nextByTime = null, monthsSinceLastRebal = null;

  if (rebalSettings.lastRebalanceDate) {
    const last = rebalSettings.lastRebalanceDate;
    nextByTime = new Date(last);
    nextByTime.setMonth(nextByTime.getMonth() + rebalSettings.freqMonths);
    monthsSinceLastRebal = (now - last) / (1000 * 60 * 60 * 24 * 30.44);
  }

  const deviationHitCount = combined.filter(x => x.needsRebalance).length;
  const timeHit = nextByTime ? (now >= nextByTime) : false;
  const deviationHit = deviationHitCount > 0;

  let firedBy = null;
  if (timeHit && deviationHit) firedBy = 'both';
  else if (timeHit) firedBy = 'time';
  else if (deviationHit) firedBy = 'deviation';

  return {
    freqMonths: rebalSettings.freqMonths,
    lastRebalanceDate: rebalSettings.lastRebalanceDate
      ? Utilities.formatDate(rebalSettings.lastRebalanceDate, 'Asia/Bangkok', 'dd/MM/yyyy') : null,
    nextByTime: nextByTime ? Utilities.formatDate(nextByTime, 'Asia/Bangkok', 'dd/MM/yyyy') : null,
    deviationPct: rebalSettings.deviationPct * 100,
    deviationHitCount,
    timeHit, deviationHit, firedBy
  };
}

// ══════════════════════════════════════════════════════════
// BUY PLAN CALCULATOR — กรอกเงินทุนรวม ระบบคำนวณแบ่งไม้ให้อัตโนมัติ
// ใช้ ATR ของหุ้นเป็นตัวกำหนดระยะห่างแต่ละไม้ (ผันผวนสูง = ไม้ห่างขึ้นเอง)
// ถ้าไม่มีข้อมูล ATR (เช่นหุ้นเพิ่ง IPO) fallback เป็น % ตายตัว — บอกวิธีที่ใช้ตรงๆ เสมอ
// ══════════════════════════════════════════════════════════

// น้ำหนัก % ต่อไม้ตามจำนวนไม้ที่เลือก (front-loaded — ไม้แรกเยอะสุด ค่อยลดลง)
const BUY_PLAN_LEG_WEIGHTS = {
  2: [60, 40],
  3: [40, 35, 25],
  4: [35, 25, 25, 15]
};

// ตัวคูณ ATR สำหรับแต่ละไม้ (ไม้แรกที่ราคาปัจจุบันเสมอ = 0×ATR)
const BUY_PLAN_ATR_STEPS = { 2: [0, 1.5], 3: [0, 1, 2], 4: [0, 0.75, 1.5, 2.5] };
// fallback % ถ้าไม่มี ATR (ห่างจากราคาปัจจุบันแบบ %)
const BUY_PLAN_PCT_STEPS = { 2: [0, 10], 3: [0, 8, 15], 4: [0, 6, 12, 20] };

function calculateBuyPlanSuggestion(ticker, market, totalBudget, numLegs, budgetCurrency, entryPrice) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    numLegs = [2, 3, 4].includes(Number(numLegs)) ? Number(numLegs) : 3;
    totalBudget = parseFloat(totalBudget);
    if (!totalBudget || totalBudget <= 0) return { success: false, error: 'กรุณาระบุเงินทุนรวมที่ถูกต้อง' };

    // ── แปลงเงินทุนถ้ากรอกเป็นบาทแต่หุ้นเป็นตลาดสหรัฐ (USD) ──
    // budgetCurrency: 'THB' | 'USD' | undefined — ถ้าไม่ระบุ ถือว่ากรอกเป็นสกุลเงินของตลาดหุ้นนั้นอยู่แล้ว
    let fxRateUsed = null;
    let originalBudgetInput = totalBudget;
    let originalCurrency = market === 'TH' ? 'THB' : 'USD';
    if (market === 'US' && budgetCurrency === 'THB') {
      fxRateUsed = getFxRate(); // THB ต่อ 1 USD
      totalBudget = totalBudget / fxRateUsed; // แปลงบาท → ดอลลาร์
      originalCurrency = 'THB';
    }

    // ── ดึงราคาสดเสมอ (ไว้โชว์เทียบ) — แต่ใช้คำนวณจริงจาก entryPrice ที่ผู้ใช้กำหนด ถ้ามีการกรอกมา
    //    เหตุผล: เวลากดคำนวณกับเวลาบันทึกแผนจริง ราคาตลาดอาจขยับไปแล้ว ผู้ใช้ควรกำหนดราคาเข้าไม้ 1 เองได้ ──
    const quote = _wlFetchYahooQuote(ticker, market); // reuse ฟังก์ชันดึงราคาจาก webapp_07_watchlist.gs
    if (!quote) return { success: false, error: 'ไม่พบราคาปัจจุบันของ ' + ticker };
    const livePrice = quote.price;

    const parsedEntryPrice = parseFloat(entryPrice);
    const currentPrice = (parsedEntryPrice && parsedEntryPrice > 0) ? parsedEntryPrice : livePrice;
    const entryPriceOverridden = Math.abs(currentPrice - livePrice) > 0.0001;

    // ── ลองหา ATR จากชีต ATR_Portfolio (ตัวเดียวกับที่ getFastSignal ใช้) ──
    let atrValue = null;
    try {
      const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
      const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);
      const rows = (market === 'TH') ? thRows : usRows;
      const found = rows.find(r => r.symbol.toUpperCase() === ticker);
      if (found && found.atr > 0) atrValue = found.atr;
    } catch (e) {
      // ไม่มีข้อมูล ATR ก็ไม่เป็นไร — fallback ไปใช้ % แทน
    }

    const weights = BUY_PLAN_LEG_WEIGHTS[numLegs];
    const usedMethod = atrValue ? 'atr' : 'percent';
    const isWholeShareOnly = (market === 'TH'); // หุ้นไทยซื้อเป็นหุ้นเต็มหน่วย ไม่มีเศษ

    const legs = weights.map((pct, i) => {
      let targetPrice;
      if (i === 0) {
        targetPrice = currentPrice; // ไม้แรก = ราคาเข้าที่กำหนด (ปัจจุบัน หรือที่ผู้ใช้แก้เอง)
      } else if (atrValue) {
        const atrMultiple = BUY_PLAN_ATR_STEPS[numLegs][i];
        targetPrice = currentPrice - (atrValue * atrMultiple);
      } else {
        const pctStep = BUY_PLAN_PCT_STEPS[numLegs][i];
        targetPrice = currentPrice * (1 - pctStep / 100);
      }
      targetPrice = Math.max(targetPrice, 0.01); // กันราคาติดลบ/ศูนย์ ในเคสหุ้นราคาต่ำมาก + ATR สูงผิดปกติ

      const budgetForLeg = totalBudget * (pct / 100);
      let estimatedShares = budgetForLeg / targetPrice;
      estimatedShares = isWholeShareOnly ? Math.floor(estimatedShares) : Math.round(estimatedShares * 10000) / 10000;

      const triggerPct = ((currentPrice - targetPrice) / currentPrice) * 100; // เก็บเป็น % เสมอ (schema ชีต BuyPlan ใช้ % ไม่ใช่ ATR multiple)

      return {
        legNumber: i + 1,
        pct,
        triggerPct: Math.round(triggerPct * 100) / 100,
        targetPrice: Math.round(targetPrice * 100) / 100,
        budgetForLeg: Math.round(budgetForLeg * 100) / 100,
        estimatedShares
      };
    });

    return {
      success: true, ticker, market,
      cur: market === 'TH' ? '฿' : '$',
      currentPrice, livePrice, entryPriceOverridden,
      totalBudget, numLegs,
      method: usedMethod, // 'atr' หรือ 'percent' — ให้ frontend บอกผู้ใช้ตรงๆ ว่าใช้วิธีไหนคำนวณ
      atrValue,
      legs,
      // ── ข้อมูลการแปลงสกุลเงิน (ถ้ามี) เพื่อความโปร่งใส ไม่ปิดบังว่าแปลงมาจากอะไร ──
      converted: !!fxRateUsed,
      fxRateUsed,
      originalBudgetInput,
      originalCurrency
    };
  } catch (e) {
    logError('calculateBuyPlanSuggestion', e);
    return { success: false, error: e.message };
  }
}

// ── บันทึกแผนที่คำนวณได้ลงชีต BuyPlan จริง (upsert — ถ้ามีแถวของ ticker/market นี้อยู่แล้วจะอัปเดตทับ) ──
function saveBuyPlanFromCalculator(ticker, market, totalBudget, legs, note) {
  try {
    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    ticker = String(ticker || '').trim().toUpperCase();

    // หา row เดิมของ ticker+market นี้ก่อน (ถ้ามี ใช้ทับแทนเพิ่มแถวใหม่ซ้ำ)
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= BUY_PLAN_SHEET.START_ROW) {
      const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
      const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 2).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toUpperCase() === ticker && String(rows[i][1]).trim().toUpperCase() === market) {
          targetRow = BUY_PLAN_SHEET.START_ROW + i;
          break;
        }
      }
    }
    if (targetRow === -1) targetRow = _getNextEmptyRow(sheet, 1, BUY_PLAN_SHEET.START_ROW);

    sheet.getRange(targetRow, 1).setValue(ticker);
    sheet.getRange(targetRow, 2).setValue(market);
    sheet.getRange(targetRow, 3).setValue('price'); // planType
    sheet.getRange(targetRow, 4).setValue(parseFloat(totalBudget));

    // เขียน leg 1-3 ลงคอลัมน์ E-J (สูงสุด 3 leg ตาม schema เดิมของชีต — ถ้าเลือก 4 ไม้ จะรวม leg 3+4 เข้าด้วยกันในช่องสุดท้าย)
    const legsToWrite = legs.slice(0, 3);
    if (legs.length > 3) {
      legsToWrite[2] = {
        pct: legs[2].pct + legs[3].pct,
        triggerPct: legs[3].triggerPct // ใช้ trigger ของไม้สุดท้ายเป็นตัวแทน
      };
    }
    legsToWrite.forEach((leg, i) => {
      sheet.getRange(targetRow, 5 + i * 2).setValue(leg.pct);
      sheet.getRange(targetRow, 6 + i * 2).setValue(leg.triggerPct);
    });

    sheet.getRange(targetRow, 13).setValue(new Date()); // วันที่เริ่มแผน
    sheet.getRange(targetRow, 14).setValue(note || 'สร้างจากเครื่องคำนวณแบ่งไม้อัตโนมัติ');

    return { success: true, row: targetRow };
  } catch (e) {
    logError('saveBuyPlanFromCalculator', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// SUB-TAB 3: แผนเข้าซื้อ — รายชื่อหุ้นให้เลือก (สำหรับ chip selector)
// ใช้รายการเดียวกับที่ต้อง rebalance ก่อน (เรียงมาก่อน) ตามด้วยตัวที่เหลือ
// ══════════════════════════════════════════════════════════
function getBuyPlanStockList() {
  try {
    const overview = getRebalanceOverviewData();
    if (!overview.success) return { success: false, error: overview.error, tickers: [] };

    const tickers = overview.stocks.map(x => ({
      ticker: x.ticker, market: x.market, needsRebalance: x.needsRebalance
    }));
    return { success: true, tickers };
  } catch (e) {
    logError('getBuyPlanStockList', e);
    return { success: false, error: e.message, tickers: [] };
  }
}

// ══════════════════════════════════════════════════════════
// SUB-TAB 3: แผนเข้าซื้อของหุ้นตัวเดียว — คำนวณสถานะแต่ละ leg/รอบ DCA
// อ่านค่าตั้งต้นจากชีต BuyPlan + เทียบกับ Transaction Log จริงว่าซื้อไปกี่ไม้แล้ว
// ══════════════════════════════════════════════════════════
function getBuyPlanForTicker(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const cfg = _readBuyPlanConfig(ticker, market);
    if (!cfg) return { success: false, error: 'ยังไม่มีแผนเข้าซื้อของ ' + ticker + ' — ไปตั้งค่าที่ชีต BuyPlan ก่อน' };

    // ── ดึงประวัติซื้อจริงของ ticker นี้จาก transaction log (นับเฉพาะประเภท 'ซื้อ') — เก็บจำนวนหุ้นด้วย (col E) ──
    const logSheetName = (market === 'TH') ? SHEETS.TH_TRANS : SHEETS.US_TRANS;
    const logRows = _getSheetDataCached(logSheetName);
    const buys = logRows
      .filter(r => String(r[2] || '').trim().toUpperCase().replace(/_C\d+$/i, '') === ticker && r[3] === 'ซื้อ')
      .map(r => ({ date: r[1], price: parseFloat(r[5]), shares: parseFloat(r[4]) }))
      .filter(x => x.date && !isNaN(x.price))
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    // ── ดึงราคาปัจจุบันสด ใช้คำนวณระยะห่างจากจุดซื้อแต่ละไม้ ──
    const quote = _wlFetchYahooQuote(ticker, market);
    const currentPrice = quote ? quote.price : null;
    const cur = market === 'TH' ? '฿' : '$';

    // ── ต้นทุนเฉลี่ยปัจจุบัน (ถ้าถืออยู่แล้ว) ใช้พรีวิวต้นทุนเฉลี่ยใหม่ต่อไม้ ──
    let holdingInfo = null;
    try {
      const h = getHoldingsData();
      const arr = (market === 'TH') ? h.th : h.us;
      holdingInfo = (arr || []).find(x => x.ticker === ticker) || null;
    } catch (e) { /* ไม่มี holdings ก็ไม่เป็นไร — ถือว่ายังไม่เคยถือ */ }

    // ── อัตราแลกเปลี่ยน (เฉพาะหุ้นสหรัฐ) ใช้โชว์ "คิดเป็น X บาท" ต่อไม้ ──
    let fxRateForDisplay = null;
    if (market === 'US') {
      try { fxRateForDisplay = getFxRate(); } catch (e) { /* ไม่มีก็ไม่โชว์ THB equivalent */ }
    }

    if (cfg.planType === 'time') {
      return { success: true, ticker, market, cur, currentPrice, planType: 'time', ...buildDcaPlanStatus(cfg, buys) };
    } else {
      return {
        success: true, ticker, market, cur, currentPrice, planType: 'price',
        ...buildPriceBasedPlanStatus(cfg, buys, currentPrice, holdingInfo, fxRateForDisplay, market)
      };
    }
  } catch (e) {
    logError('getBuyPlanForTicker', e);
    return { success: false, error: e.message };
  }
}

// ── อ่าน config แผนซื้อจากชีต BuyPlan ──
function _readBuyPlanConfig(ticker, market) {
  const sheet = getSheet(BUY_PLAN_SHEET.NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < BUY_PLAN_SHEET.START_ROW) return null;

  const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
  const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 15).getValues(); // ← 14→15

  for (let i = 0; i < rows.length; i++) {
    const rTicker = String(rows[i][0] || '').trim().toUpperCase();
    const rMarket = String(rows[i][1] || '').trim().toUpperCase();
    if (rTicker === ticker && (!market || rMarket === market)) {
      return {
        ticker: rTicker, market: rMarket,
        planType: String(rows[i][2] || 'price').trim().toLowerCase(),
        budget: parseFloat(rows[i][3]) || 0,
        legs: [
          { pct: parseFloat(rows[i][4]) || 0, triggerPct: parseFloat(rows[i][5]) },
          { pct: parseFloat(rows[i][6]) || 0, triggerPct: parseFloat(rows[i][7]) },
          { pct: parseFloat(rows[i][8]) || 0, triggerPct: parseFloat(rows[i][9]) }
        ].filter(l => l.pct > 0),
        dcaFreqDays: parseFloat(rows[i][10]) || null,
        dcaAmount: parseFloat(rows[i][11]) || null,
        startDate: rows[i][12] instanceof Date ? rows[i][12] : null,
        note: rows[i][13] || '',
        referencePrice: parseFloat(rows[i][14]) || null // ← ใหม่
      };
    }
  }
  return null;
}


// ── คำนวณสถานะแผน Price-based: จับคู่ไม้ซื้อจริงกับแต่ละ leg ตาม "ราคา trigger จริง"
//    ไม่ใช่แค่นับจำนวนไม้ที่ซื้อ — เพราะถ้าคนซื้อไม้พิเศษนอกแผน การนับแบบเดิมจะรายงาน
//    สถานะผิด (เช่น บอกว่า leg 2 เสร็จแล้ว ทั้งที่จริงคือซื้อไม้อื่นที่ไม่ตรงเงื่อนไข)
//    วิธีนี้: leg จะถือว่า "done" ก็ต่อเมื่อมีไม้ซื้อจริงที่ราคา ≤ targetPrice ของ leg นั้น
//    และไม้นั้นต้องเกิดขึ้น "หลัง" ไม้ที่จับคู่กับ leg ก่อนหน้าแล้วเท่านั้น (รักษาลำดับเวลา) ──
function buildPriceBasedPlanStatus(cfg, buys, currentPrice, holdingInfo, fxRateForDisplay, market) {
  const isWholeShareOnly = (market === 'TH');
  const firstEntryPrice = buys.length ? buys[0].price : cfg.referencePrice; // ← fallback ใหม่
  const usingReferencePrice = !buys.length && !!cfg.referencePrice;

  let searchFromIdx = 0;
  const steps = cfg.legs.map((leg, i) => {
    if (i === 0) {
      const executed = buys.length > 0;
      const result = {
        legNumber: 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice: firstEntryPrice,
        status: executed ? 'done' : 'not_yet', // ← เปลี่ยนจาก 'pending' → เข้า flow canBuyNow เหมือนไม้อื่น
        executedPrice: executed ? buys[0].price : null,
        executedShares: executed && !isNaN(buys[0].shares) ? buys[0].shares : null,
        executedDate: executed ? Utilities.formatDate(new Date(buys[0].date), 'Asia/Bangkok', 'dd/MM/yyyy') : null,
        matchNote: usingReferencePrice ? 'อ้างอิงราคา ณ ตอนสร้างแผน (ยังไม่เคยถือ)' : null
      };
      if (executed) searchFromIdx = 1;
      return result;
    }

    if (firstEntryPrice === null) {
      return { legNumber: i + 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice: null,
        status: 'pending', executedPrice: null, executedShares: null, executedDate: null, matchNote: 'ไม่มีราคาอ้างอิง' };
    }

    const targetPrice = firstEntryPrice * (1 - Math.abs(leg.triggerPct) / 100);
    let matchedIdx = -1;
    for (let b = searchFromIdx; b < buys.length; b++) {
      if (buys[b].price <= targetPrice) { matchedIdx = b; break; }
    }

    if (matchedIdx !== -1) {
      searchFromIdx = matchedIdx + 1;
      return {
        legNumber: i + 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice,
        status: 'done', executedPrice: buys[matchedIdx].price,
        executedShares: !isNaN(buys[matchedIdx].shares) ? buys[matchedIdx].shares : null,
        executedDate: Utilities.formatDate(new Date(buys[matchedIdx].date), 'Asia/Bangkok', 'dd/MM/yyyy'),
        matchNote: null
      };
    }
    return { legNumber: i + 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice,
      status: 'not_yet', executedPrice: null, executedShares: null, executedDate: null, matchNote: null };
  });

  let nextAssigned = false;
  steps.forEach(s => {
    if (s.status === 'not_yet') { s.status = nextAssigned ? 'pending' : 'next'; nextAssigned = true; }
  });

  if (currentPrice) {
    steps.forEach(s => {
      if (s.status === 'done' || s.targetPrice === null) return;
      s.distanceAmount = currentPrice - s.targetPrice;
      s.distancePct = (s.distanceAmount / currentPrice) * 100;
      s.canBuyNow = currentPrice <= s.targetPrice; // ← ตอนนี้ครอบคลุมไม้ 1 ด้วย
    });
  }

  steps.forEach(s => {
    if (s.status === 'done') {
      s.actualBudgetSpent = s.executedShares !== null ? s.executedPrice * s.executedShares : cfg.budget * (s.pct / 100);
      s.budgetTHBEquivalent = (market === 'US' && fxRateForDisplay) ? s.actualBudgetSpent * fxRateForDisplay : null;
    } else if (s.targetPrice !== null) {
      const budgetForLeg = cfg.budget * (s.pct / 100);
      let estimatedShares = budgetForLeg / s.targetPrice;
      estimatedShares = isWholeShareOnly ? Math.floor(estimatedShares) : Math.round(estimatedShares * 10000) / 10000;
      s.budgetForLeg = budgetForLeg;
      s.estimatedShares = estimatedShares;
      s.budgetTHBEquivalent = (market === 'US' && fxRateForDisplay) ? budgetForLeg * fxRateForDisplay : null;

      const curShares = holdingInfo ? (parseFloat(holdingInfo.sharesRemain) || 0) : 0;
      const curTotalCost = holdingInfo ? (parseFloat(holdingInfo.totalCost) || 0) : 0;
      const curAvgCost = holdingInfo ? (parseFloat(holdingInfo.avgCost) || null) : null;
      const newTotalCost = curTotalCost + (s.targetPrice * estimatedShares);
      const newTotalShares = curShares + estimatedShares;
      const newAvgCost = newTotalShares > 0 ? newTotalCost / newTotalShares : s.targetPrice;

      s.curAvgCost = curAvgCost;
      s.newAvgCost = newAvgCost;
      s.avgCostChangeAmount = (curAvgCost !== null) ? (newAvgCost - curAvgCost) : null;
      s.avgCostChangePct = curAvgCost ? ((newAvgCost - curAvgCost) / curAvgCost) * 100 : null;
    }
  });

  return { budget: cfg.budget, firstEntryPrice, usingReferencePrice, steps, note: cfg.note };
}

// ── คำนวณสถานะแผน Time-based (DCA): รอบถัดไปคือเมื่อไหร่ ──
function buildDcaPlanStatus(cfg, buys) {
  const lastBuyDate = buys.length ? new Date(buys[buys.length - 1].date) : cfg.startDate;
  let nextDcaDate = null, daysUntilNext = null;

  if (lastBuyDate && cfg.dcaFreqDays) {
    nextDcaDate = new Date(lastBuyDate);
    nextDcaDate.setDate(nextDcaDate.getDate() + cfg.dcaFreqDays);
    daysUntilNext = Math.ceil((nextDcaDate - new Date()) / (1000 * 60 * 60 * 24));
  }

  return {
    budget: cfg.budget,
    dcaFreqDays: cfg.dcaFreqDays,
    dcaAmount: cfg.dcaAmount,
    totalBuysExecuted: buys.length,
    lastBuyDate: lastBuyDate ? Utilities.formatDate(lastBuyDate, 'Asia/Bangkok', 'dd/MM/yyyy') : null,
    nextDcaDate: nextDcaDate ? Utilities.formatDate(nextDcaDate, 'Asia/Bangkok', 'dd/MM/yyyy') : null,
    daysUntilNext,
    dueNow: daysUntilNext !== null && daysUntilNext <= 0,
    note: cfg.note
  };
}

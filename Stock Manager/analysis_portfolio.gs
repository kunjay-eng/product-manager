// ============================================================
// analysis_portfolio.gs — เฟส 3: Analysis_Portfolio
// วิเคราะห์สัดส่วนพอร์ต (Weight) + Rebalance + Concentration Risk
// ครอบคลุมหุ้น (US+TH) และกองทุน รวมกันเป็นสกุลเงินเดียว (THB)
//
// ต่อยอดจากเฟส 1 (Settings + StockMode) และใช้ getHoldings()/
// getFundHoldings()/getFxRate() ที่มีอยู่แล้วในโปรเจกต์ ไม่ duplicate
// การอ่านชีตใดๆ
//
// หมายเหตุ: สูตรต้นแบบ (Analysis_Portfolio ในสเปรดชีตต้นแบบ) อ้างอิง
// Settings!E12 เป็นเกณฑ์ % ส่วนต่างที่ต้อง Rebalance ซึ่งเป็นเซลล์ใหม่ที่
// ไม่มีในชีต Settings ที่สร้างไว้ตอนเฟส 1 — ไฟล์นี้จึงมีฟังก์ชัน
// setupRebalanceThreshold() เพิ่มเซลล์นี้ให้ (รันครั้งเดียว ปลอดภัย
// ไม่ทับถ้ามีอยู่แล้ว) โดยไม่ต้องแก้ settings_stockmode.gs เดิม
// ============================================================

const REBALANCE_THRESHOLD_CELL       = 'E11'; // % ส่วนต่างน้ำหนักที่ต้อง Rebalance
const REBALANCE_THRESHOLD_LABEL_CELL = 'B11';

// ──────────────────────────────────────
// รันครั้งเดียว (ถ้ายังไม่เคยรัน) เพื่อเพิ่มเซลล์ Rebalance Threshold
// เข้าไปในชีต Settings ที่สร้างไว้แล้วตั้งแต่เฟส 1
// ──────────────────────────────────────
function setupRebalanceThreshold() {
  try {
    const sheet = getSheet(SETTINGS_SHEET.SHEET);
    const existing = sheet.getRange(REBALANCE_THRESHOLD_CELL).getValue();
    if (existing !== '' && existing !== null) {
      try { SpreadsheetApp.getUi().alert('มีค่า Rebalance Threshold อยู่แล้ว ไม่ได้แก้ไข'); } catch (e2) {}
      return;
    }
    sheet.getRange(REBALANCE_THRESHOLD_LABEL_CELL)
      .setValue('Rebalance Threshold — ส่วนต่างน้ำหนักที่ต้องปรับพอร์ต (%)');
    sheet.getRange(REBALANCE_THRESHOLD_CELL).setValue(0.05).setNumberFormat('0.00%');
    sheet.getRange(REBALANCE_THRESHOLD_CELL).setBackground('#fff2cc');
    try { SpreadsheetApp.getUi().alert('✅ เพิ่มค่า Rebalance Threshold ในชีต Settings แล้ว (default 5%)'); } catch (e2) {}
  } catch (e) {
    logError('setupRebalanceThreshold', e);
    try { SpreadsheetApp.getUi().alert('❌ ' + e.message + ' — ต้องรัน setupNewSheets() (เฟส 1) ก่อน'); } catch (e2) {}
  }
}

function getRebalanceThresholdPct() {
  try {
    const sheet = getSheet(SETTINGS_SHEET.SHEET);
    const v = Number(sheet.getRange(REBALANCE_THRESHOLD_CELL).getValue());
    return v || 0.05;
  } catch (e) {
    return 0.05;
  }
}

// ══════════════════════════════════════════════════════════
// วิเคราะห์สัดส่วนพอร์ตทั้งหมด (หุ้น US+TH+กองทุน รวมเป็น THB)
// เรียกจากหน้าเว็บผ่าน google.script.run.getPortfolioAnalysis()
// ══════════════════════════════════════════════════════════
function getPortfolioAnalysis() {
  try {
    const fxRate = getFxRate();
    const modeMap = getStockModeMap();
    const rebalanceThresholdPct = getRebalanceThresholdPct();
    const riskSettings = getRiskSettings();

    const usHold   = getHoldings(SHEETS.US_HOLD);
    const thHold   = getHoldings(SHEETS.TH_HOLD);
    const fundHold = getFundHoldings();

    const rows = [];

    usHold.forEach(h => {
      if (h.valueNow <= 0) return;
      rows.push({ ticker: h.ticker, market: 'สหรัฐ', assetType: 'หุ้น',
                  valueNative: h.valueNow, valueTHB: h.valueNow * fxRate });
    });
    thHold.forEach(h => {
      if (h.valueNow <= 0) return;
      rows.push({ ticker: h.ticker, market: 'ไทย', assetType: 'หุ้น',
                  valueNative: h.valueNow, valueTHB: h.valueNow });
    });
    fundHold.forEach(h => {
      if (h.valueNow <= 0) return;
      rows.push({ ticker: h.name, market: 'ไทย', assetType: 'กองทุน',
                  valueNative: h.valueNow, valueTHB: h.valueNow });
    });

    const totalTHB = rows.reduce((s, r) => s + r.valueTHB, 0);
    const countPositions = rows.length;

    rows.forEach(r => {
      const cfg = modeMap[String(r.ticker).toUpperCase()] || {};
      r.trendGroup = cfg.trendGroup || 'ไม่ระบุ';
      r.currentWeightPct = totalTHB > 0 ? (r.valueTHB / totalTHB) * 100 : 0;

      // Target Weight: ใช้ override ก่อน > auto ที่ตั้งไว้ > equal-weight fallback
      // (ตรงตามสูตรต้นแบบ: 1/COUNTA ของหุ้นทั้งหมด ถ้าไม่ได้ตั้งอะไรเลย)
      let targetWeightPct;
      if (cfg.targetWeightOverride !== null && cfg.targetWeightOverride !== undefined) {
        targetWeightPct = cfg.targetWeightOverride * 100;
      } else if (cfg.targetWeightAuto !== null && cfg.targetWeightAuto !== undefined) {
        targetWeightPct = cfg.targetWeightAuto * 100;
      } else {
        targetWeightPct = countPositions > 0 ? (100 / countPositions) : 0;
      }
      r.targetWeightPct = targetWeightPct;
      r.diffPct = r.currentWeightPct - targetWeightPct;
      r.needsRebalance = Math.abs(r.diffPct) >= (rebalanceThresholdPct * 100);
      r.overConcentration = r.currentWeightPct >= (riskSettings.concentrationWarnPct * 100);
    });

    rows.sort((a, b) => b.valueTHB - a.valueTHB);

    // สรุปตามกลุ่ม Trend Group (Growth/Value/Defensive/Dividend)
    const groupMap = {};
    rows.forEach(r => {
      const g = r.trendGroup || 'ไม่ระบุ';
      if (!groupMap[g]) groupMap[g] = { group: g, valueTHB: 0 };
      groupMap[g].valueTHB += r.valueTHB;
    });
    const groups = Object.values(groupMap).map(g => ({
      group: g.group, valueTHB: g.valueTHB,
      currentWeightPct: totalTHB > 0 ? (g.valueTHB / totalTHB) * 100 : 0
    })).sort((a, b) => b.valueTHB - a.valueTHB);

    return {
      success: true,
      totalTHB: totalTHB,
      rebalanceThresholdPct: rebalanceThresholdPct * 100,
      concentrationWarnPct: riskSettings.concentrationWarnPct * 100,
      rows: rows,
      groups: groups,
      needRebalanceCount: rows.filter(r => r.needsRebalance).length,
      overConcentrationCount: rows.filter(r => r.overConcentration).length
    };
  } catch (e) {
    logError('getPortfolioAnalysis', e);
    return { success: false, error: e.message, rows: [], groups: [],
             totalTHB: 0, needRebalanceCount: 0, overConcentrationCount: 0 };
  }
}

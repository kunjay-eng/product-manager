// ============================================================
// summary.gs — เฟส 3: หน้า Summary (หน้าแรกใหม่)
// รวมมูลค่าพอร์ต + breakdown กำไร/ขาดทุนแบบเต็ม (กำไรทิพย์/กำไรจริง/
// ปันผล/ค่าธรรมเนียม/กำไรค่าเงิน) + Financial Goal + Allocation + Alerts
// + Top Movers ตามต้นแบบที่ให้มา (Stock_Portfolio_v1.docx)
//
// ใช้ฟังก์ชันที่มีอยู่แล้วในโปรเจกต์ทั้งหมด ไม่คำนวณซ้ำ:
//   collectAllData()      (data.gs)   — unrealized/realized/dividend/cash/holdings/top movers
//   getFinancialGoalData()(Code.gs)   — เป้าหมาย + CAGR (เรียกแยกจากหน้าเว็บ ไม่ได้ผูกในนี้)
//   getRealizedPnL()      (data.gs)
//   _getFeeTH()           (summary_dashboard.gs)
//   _getATRRows()         (atr_report.gs) — ใช้เช็คหุ้นที่ใกล้/หลุด Cut Stop
//   getStockModeMap()     (settings_stockmode.gs) — target weight (rebalance เฟส 1)
// ============================================================

// ค่าธรรมเนียมฝั่งสหรัฐ (คอลัมน์เดียวกับที่ sendSummaryDashboard() อ่านอยู่แล้ว
// ที่ US_REAL!O5 — แยกเป็นฟังก์ชันเผื่อไฟล์อื่นเรียกใช้ซ้ำได้)
function _getFeeUS() {
  try {
    const sheet = getSheet(SHEETS.US_REAL);
    return Number(sheet.getRange('O5').getValue()) || 0;
  } catch (e) {
    logError('_getFeeUS', e);
    return 0;
  }
}

// ──────────────────────────────────────
// เรียกจากหน้าเว็บผ่าน google.script.run.getSummaryData()
// ──────────────────────────────────────
/** function getSummaryData() {
  try {
    const d = collectAllData(); // ของเดิมทั้งหมด: cash, holdings, unrealized, realized, dividend, top movers, fxRate
    const fx = d.fxRate;

    // ── เงินสด แยกสกุลเงิน (สำหรับแท็บ "🏦 เงินสด" ใหม่) ──
    const cashBal = getCashBalances(); // { thb, usd, total, deposit, withdraw, totalCost }

    // ── ค่าธรรมเนียมรวม (สูตรเดียวกับ sendSummaryDashboard) ──
    const feeUS = _getFeeUS();
    const feeTH = _getFeeTH();
    const feeTotalTHB = feeTH + feeUS * fx;

    // ── กำไร/ขาดทุนค่าเงิน (ประมาณการ) ──
    // เทียบมูลค่า USD ที่เคยฝากเข้ามาทั้งหมด (F6) คิดที่อัตราแลกเปลี่ยนปัจจุบัน
    // กับจำนวนเงินบาทที่จ่ายจริงตอนแลกซื้อ USD (C11) — ส่วนต่างคือกำไร/ขาดทุน
    // จากอัตราแลกเปลี่ยนล้วนๆ แยกออกจากกำไร/ขาดทุนจากราคาหุ้น
    // (เป็นค่าประมาณการ เพราะไม่ได้เก็บ FX เฉลี่ยถ่วงน้ำหนักจริงทุกรายการ)
    const cashSheet    = getSheet(SHEETS.CASH);
    const exchangeTHB  = Number(cashSheet.getRange('C11').getValue()) || 0;
    const totalUSDIn   = Number(cashSheet.getRange('F6').getValue())  || 0;
    const fxGainLossTHB = (totalUSDIn * fx) - exchangeTHB;

    // ── รวมสุทธิ ──
    const netTotalTHB = d.totalUnrealTHB + d.totalRealTHB + d.divTotalTHB
                         - feeTotalTHB + fxGainLossTHB;

    // ── Allocation: เงินสด / หุ้น / กองทุน ──
    const allocCash  = d.cashTotal;
    const allocStock = d.thPortTHB + d.usPortTHB;
    const allocFund  = d.fundPortTHB;
    const allocTotal = allocCash + allocStock + allocFund;

    // ── Alert: ใกล้/หลุด Cut Stop (ใช้ข้อมูล ATR เดิม) ──
    const usATR = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thATR = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);
    const cutStopCount = [...usATR, ...thATR]
      .filter(r => Number(r.stopDistance) <= 5) // ≤5% ถือว่าใกล้/หลุดแล้ว
      .length;

    // ── Alert: ต้อง Rebalance (เทียบ target weight ที่ตั้งไว้ใน StockMode เฟส 1)
    //    นับเฉพาะตัวที่ตั้ง Override ไว้จริง — ถ้ายังไม่มีใครตั้งจะได้ 0
    //    (การคำนวณ Auto weight เต็มรูปแบบรอเฟส Analysis_Portfolio) ──
    let rebalanceCount = 0;
    try {
      const modeMap = getStockModeMap();
      const allHold = [...d.usHoldings, ...d.thHoldings];
      const totalPortTHB = d.totalPortTHB;
      allHold.forEach(h => {
        const cfg = modeMap[String(h.ticker).toUpperCase()];
        if (!cfg || cfg.targetWeightOverride === null || cfg.targetWeightOverride === undefined) return;
        const currentWeight = totalPortTHB > 0 ? (h.valueNow / totalPortTHB) * 100 : 0;
        if (Math.abs(currentWeight - cfg.targetWeightOverride) > 5) rebalanceCount++;
      });
    } catch (e) { /* ยังไม่ได้รัน setupNewSheets() ของเฟส 1 ก็ไม่เป็นไร ปล่อยเป็น 0 

    const totalAssetTHB = d.totalPortTHB + d.cashTotal;

    return {
      success: true,

      totalAssetTHB: totalAssetTHB,
      netTotalTHB: netTotalTHB,

      unrealizedTHB: d.totalUnrealTHB,
      realizedTHB: d.totalRealTHB,
      dividendTHB: d.divTotalTHB,
      feeTotalTHB: feeTotalTHB,
      fxGainLossTHB: fxGainLossTHB,

      allocation: { cash: allocCash, stock: allocStock, fund: allocFund, total: allocTotal },

      cutStopCount: cutStopCount,
      rebalanceCount: rebalanceCount,

      usTop3: d.usTop3, usWorst: d.usWorst,
      thTop3: d.thTop3, thWorst: d.thWorst,

      th:   { valueTHB: d.thPortTHB,  unrealTHB: d.thUnrealTHB, realTHB: d.thRealTHB },
      us:   { valueUSD: d.usPortUSD,  valueTHB: d.usPortTHB, unrealUSD: d.usUnrealUSD, unrealTHB: d.usUnrealTHB, realUSD: d.usRealUSD, realTHB: d.usRealTHB },
      fund: { valueTHB: d.fundPortTHB, unrealTHB: d.fundUnrealTHB },
      cash: { thb: cashBal.thb, usd: cashBal.usd, usdInTHB: cashBal.usd * fx, totalTHB: cashBal.total },

      netInvest: d.netInvest,
      fxRate: fx,
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
    };
  } catch (e) {
    logError('getSummaryData', e);
    return { success: false, error: e.message };
  }
}*/


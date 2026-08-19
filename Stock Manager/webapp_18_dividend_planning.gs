// ══════════════════════════════════════════════════════════
// webapp_18_dividend_planning.gs — แผนเงินปันผลทั้งปี + โปรเจกชัน
// ------------------------------------------------------------
// ใช้ trailing 12 เดือนจริง (ไม่ annualize จาก YTD) เพื่อความแม่นยำ
// รวมทั้ง US+TH (ยังไม่รวมกองทุน เพราะระบบยังไม่มี field ปันผลกองทุนแยก)
// ══════════════════════════════════════════════════════════

function getDividendPlanningData() {
  try {
    const goal = getFinancialGoalData();
    if (!goal.success) return { success: false, error: goal.error };

    const divSheet = getSheet(SHEETS.DIV);
    const annualTarget = Number(divSheet.getRange('AA7').getValue()) || 0;
    const ytd = getDividendYTD(); // { thTHB, usUSD, totalTHB, target }

    // ── รวมปันผล 12 เดือนล่าสุด (ไม่ใช่ปีปฏิทิน) จาก transaction log ปันผลจริง ──
    const trailing12mo = _getTrailing12MonthDividendTHB();

    const totalAssetTHB = goal.totalAsset;
    const currentYieldPct = totalAssetTHB > 0 ? (trailing12mo.totalTHB / totalAssetTHB) * 100 : 0;

    const progressPct = annualTarget > 0 ? (ytd.totalTHB / annualTarget) * 100 : null;
    const remainingToTarget = annualTarget > 0 ? Math.max(0, annualTarget - ytd.totalTHB) : null;

    // ── เดือนที่เหลือในปีนี้ ใช้ประมาณว่าถ้าอัตราปัจจุบันคงที่ จะจบปีที่เท่าไหร่ ──
    const now = new Date();
    const monthsElapsed = now.getMonth() + 1;
    const monthsRemaining = 12 - monthsElapsed;
    const avgMonthlyThisYear = monthsElapsed > 0 ? ytd.totalTHB / monthsElapsed : 0;
    const projectedYearEndTHB = ytd.totalTHB + (avgMonthlyThisYear * monthsRemaining);

    return {
      success: true,
      annualTarget,
      ytdDividendTHB: ytd.totalTHB,
      progressPct,
      remainingToTarget,
      projectedYearEndTHB: _taxRound(projectedYearEndTHB, 2),
      trailing12moTHB: _taxRound(trailing12mo.totalTHB, 2),
      trailing12moByMarket: { th: _taxRound(trailing12mo.thTHB, 2), us: _taxRound(trailing12mo.usTHB, 2) },
      totalAssetTHB,
      currentYieldPct: _taxRound(currentYieldPct, 3),
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('getDividendPlanningData', e);
    return { success: false, error: e.message };
  }
}

// ── รวมปันผลสุทธิ (netTHB) ของ 12 เดือนล่าสุดนับถอยจากวันนี้ (ไม่ใช่ปีปฏิทิน)
//    อ่านจากชีตปันผลตรงๆ ด้วย DIV_COL เดิม (ตัวเดียวกับ getDividendMonthly ใน data.gs) ──
function _getTrailing12MonthDividendTHB() {
  const sheet = getSheet(SHEETS.DIV);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.DIV) return { totalTHB: 0, thTHB: 0, usTHB: 0 };

  const numRows = lastRow - START_ROW.DIV + 1;
  const rows = sheet.getRange(START_ROW.DIV, 1, numRows, 17).getValues();

  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 12);

  let totalTHB = 0, thTHB = 0, usTHB = 0;
  rows.forEach(row => {
    const payDate = row[DIV_COL.PAY_DATE - 1];
    const market = String(row[DIV_COL.MARKET - 1] || '').trim().toUpperCase();
    const netTHB = Number(row[DIV_COL.NET_THB - 1]) || 0;
    if (!(payDate instanceof Date) || payDate < cutoff) return;

    totalTHB += netTHB;
    if (market === 'TH') thTHB += netTHB; else usTHB += netTHB;
  });

  return { totalTHB, thTHB, usTHB };
}

// ── บันทึกเป้าหมายปันผลทั้งปีใหม่ — เขียนกลับไปที่ AA7 (cell เดิมที่ getDividendYTD() อ่านอยู่แล้ว) ──
function setDividendAnnualTarget(amount) {
  try {
    const target = parseFloat(amount);
    if (isNaN(target) || target < 0) return { success: false, error: 'กรุณาระบุจำนวนเงินที่ถูกต้อง' };

    const sheet = getSheet(SHEETS.DIV);
    sheet.getRange('AA7').setValue(target).setNumberFormat('#,##0.00');
    return { success: true, target };
  } catch (e) {
    logError('setDividendAnnualTarget', e);
    return { success: false, error: e.message };
  }
}


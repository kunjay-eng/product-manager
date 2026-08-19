// ============================================================
// webapp_16_tax_report.gs — รายงานกำไร/ขาดทุนหุ้นสหรัฐสำหรับยื่นภาษี
// ------------------------------------------------------------
// คำนวณ Realized Gain/Loss รายปี จากประวัติธุรกรรม US_TRANS ทั้งหมด
// ด้วยวิธี Average Cost (สอดคล้องกับ avgCost ที่ใช้ในหน้า Holdings)
// โดยจำลอง (simulate) ธุรกรรมทั้งหมดตามลำดับเวลาใหม่ทั้งหมด — ไม่พึ่งพา
// ยอดสะสมใน US_REAL sheet เพราะยอดนั้นเป็น summary รวม ไม่ใช่ต่อรายการ
//
// ⚠️ ข้อจำกัด/สมมติฐานที่ต้องรู้ก่อนใช้ยื่นภาษีจริง:
// 1. ใช้ Average Cost ไม่ใช่ FIFO/Specific Lot — ถ้าต้องการ FIFO ต้องคำนวณแยก
// 2. ต้นทุนรวมค่าคอมมิชชัน/TAF/VAT ฝั่งซื้อ, proceeds หักค่าธรรมเนียมฝั่งขายแล้ว
// 3. แปลง USD→THB ด้วยอัตราแลกเปลี่ยนที่บันทึกไว้ ณ วันที่ขายจริง (ไม่ใช่เรทเฉลี่ย
//    ถ่วงน้ำหนักของทุกไม้ที่ขาย) — ถ้าแถวไหนไม่มีเรทบันทึกไว้จะ fallback เรทปัจจุบัน
//    และ flag rateIsEstimated:true
// 4. ภาษีเงินได้บุคคลธรรมดาไทยสำหรับกำไรหุ้นต่างประเทศ ขึ้นกับปีที่ "นำเงินกลับเข้าไทย"
//    (remittance) ไม่ใช่ปีที่ขาย — รายงานนี้จัดกลุ่มตาม "ปีที่ขาย" เท่านั้น ต้องเช็คปีที่
//    โอนเงินกลับเองเพิ่มเติม
// 5. ไม่ใช่คำแนะนำทางภาษี ควรตรวจสอบกับผู้เชี่ยวชาญด้านภาษีก่อนยื่นจริง
// ============================================================

function getUSCapitalGainTaxReport(year) {
  try {
    year = parseInt(year, 10) || new Date().getFullYear();
    const sheet = getSheet(SHEETS.US_TRANS);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.HOLD) {
      return { success: true, year, sales: [], byTicker: [], summary: _taxEmptySummary() };
    }

    const numRows = lastRow - START_ROW.HOLD + 1;
    const rows = sheet.getRange(START_ROW.HOLD, 1, numRows, 17).getValues();

    // ── แปลงเป็น object + กรองแถวที่ใช้งานได้ ──
    const txns = rows.map((r, i) => ({
      rowIndex: i,
      date: r[1] instanceof Date ? r[1] : (r[1] ? new Date(r[1]) : null),
      tickerRaw: String(r[2] || '').trim().toUpperCase(),
      type: String(r[3] || '').trim(),
      shares: Number(r[4]) || 0,
      price: Number(r[5]) || 0,
      commission: Number(r[7]) || 0,
      taf: Number(r[8]) || 0,
      tafFee: Number(r[9]) || 0,
      vat: Number(r[10]) || 0,
      exchangeRate: Number(r[16]) || 0
    })).filter(t => t.date && t.tickerRaw && t.shares > 0 && (t.type === 'ซื้อ' || t.type === 'ขาย'));

    // เรียงตามวันที่ (เก่า→ใหม่) ถ้าวันเดียวกันให้เรียงตามลำดับแถวเดิม
    txns.sort((a, b) => a.date - b.date || a.rowIndex - b.rowIndex);

    const lots = {}; // tickerRaw -> { shares, costUSD }
    const sales = [];
    let missingRateCount = 0;

    txns.forEach(t => {
      if (!lots[t.tickerRaw]) lots[t.tickerRaw] = { shares: 0, costUSD: 0 };
      const lot = lots[t.tickerRaw];
      const feesTotal = t.commission + t.taf + t.tafFee + t.vat;

      if (t.type === 'ซื้อ') {
        lot.shares += t.shares;
        lot.costUSD += (t.shares * t.price) + feesTotal;
        return;
      }

      // ── ขาย ──
      if (lot.shares <= 0) return; // ไม่มีสถานะให้ขาย (ข้อมูลตกหล่น) ข้ามไปกันพัง ไม่ fabricate
      const avgCostPerShare = lot.costUSD / lot.shares;
      const sellShares = Math.min(t.shares, lot.shares); // กันเคสขายเกินจากข้อมูลไม่ครบ
      const costBasisUSD = avgCostPerShare * sellShares;
      const proceedsUSD = (sellShares * t.price) - feesTotal;
      const gainUSD = proceedsUSD - costBasisUSD;

      lot.shares -= sellShares;
      lot.costUSD -= costBasisUSD;
      if (lot.shares < 0.0000001) { lot.shares = 0; lot.costUSD = 0; } // เก็บกวาด floating point

      const saleYear = parseInt(Utilities.formatDate(t.date, 'Asia/Bangkok', 'yyyy'), 10);
      if (saleYear !== year) return; // จำลองครบทุกปีเพื่อความถูกต้องของ avgCost แต่เก็บผลลัพธ์เฉพาะปีที่ขอ

      const rateIsEstimated = !(t.exchangeRate > 0);
      const rateUsed = rateIsEstimated ? getFxRate() : t.exchangeRate;
      if (rateIsEstimated) missingRateCount++;

      sales.push({
        date: Utilities.formatDate(t.date, 'Asia/Bangkok', 'dd/MM/yyyy'),
        dateSort: t.date.getTime(),
        ticker: t.tickerRaw.replace(/_C\d+$/i, ''),
        tickerRaw: t.tickerRaw,
        sharesSold: _taxRound(sellShares, 7),
        avgCostPerShare: _taxRound(avgCostPerShare, 4),
        sellPrice: t.price,
        feesUSD: _taxRound(feesTotal, 2),
        costBasisUSD: _taxRound(costBasisUSD, 2),
        proceedsUSD: _taxRound(proceedsUSD, 2),
        gainUSD: _taxRound(gainUSD, 2),
        exchangeRateUsed: rateUsed,
        rateIsEstimated,
        gainTHB: _taxRound(gainUSD * rateUsed, 2)
      });
    });

    sales.sort((a, b) => b.dateSort - a.dateSort); // ล่าสุดก่อน
    sales.forEach(s => delete s.dateSort);

    // ── สรุปตามหุ้น ──
    const byTickerMap = {};
    sales.forEach(s => {
      if (!byTickerMap[s.ticker]) byTickerMap[s.ticker] = { ticker: s.ticker, gainUSD: 0, gainTHB: 0, salesCount: 0 };
      byTickerMap[s.ticker].gainUSD += s.gainUSD;
      byTickerMap[s.ticker].gainTHB += s.gainTHB;
      byTickerMap[s.ticker].salesCount++;
    });
    const byTicker = Object.values(byTickerMap)
      .map(x => ({ ...x, gainUSD: _taxRound(x.gainUSD, 2), gainTHB: _taxRound(x.gainTHB, 2) }))
      .sort((a, b) => b.gainTHB - a.gainTHB);

    const gains = sales.filter(s => s.gainTHB > 0);
    const losses = sales.filter(s => s.gainTHB < 0);
    const totalGainUSD = gains.reduce((s, x) => s + x.gainUSD, 0);
    const totalLossUSD = losses.reduce((s, x) => s + x.gainUSD, 0);
    const totalGainTHB = gains.reduce((s, x) => s + x.gainTHB, 0);
    const totalLossTHB = losses.reduce((s, x) => s + x.gainTHB, 0);

    return {
      success: true,
      year,
      sales,
      byTicker,
      summary: {
        totalSales: sales.length,
        totalGainUSD: _taxRound(totalGainUSD, 2),
        totalLossUSD: _taxRound(totalLossUSD, 2),
        netGainUSD: _taxRound(totalGainUSD + totalLossUSD, 2),
        totalGainTHB: _taxRound(totalGainTHB, 2),
        totalLossTHB: _taxRound(totalLossTHB, 2),
        netGainTHB: _taxRound(totalGainTHB + totalLossTHB, 2),
        missingRateCount
      },
      method: 'average_cost',
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('getUSCapitalGainTaxReport', e);
    return { success: false, error: e.message };
  }
}

// ── รายชื่อปีที่มีรายการขายจริง — ใช้ทำ dropdown เลือกปี (แพตเทิร์นเดียวกับ getDividendYears()) ──
function getUSTaxReportYears() {
  try {
    const sheet = getSheet(SHEETS.US_TRANS);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.HOLD) return [new Date().getFullYear()];

    const numRows = lastRow - START_ROW.HOLD + 1;
    const rows = sheet.getRange(START_ROW.HOLD, 1, numRows, 4).getValues();
    const years = new Set();
    rows.forEach(r => {
      const date = r[1];
      const type = String(r[3] || '').trim();
      if (date instanceof Date && type === 'ขาย') {
        years.add(parseInt(Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy'), 10));
      }
    });
    if (!years.size) years.add(new Date().getFullYear());
    return Array.from(years).sort((a, b) => b - a);
  } catch (e) {
    logError('getUSTaxReportYears', e);
    return [new Date().getFullYear()];
  }
}

function _taxEmptySummary() {
  return { totalSales: 0, totalGainUSD: 0, totalLossUSD: 0, netGainUSD: 0, totalGainTHB: 0, totalLossTHB: 0, netGainTHB: 0, missingRateCount: 0 };
}

function _taxRound(v, d) {
  const f = Math.pow(10, d || 0);
  return Math.round((v + Number.EPSILON) * f) / f;
}

// ── ทดสอบใน Apps Script Editor ──
function testUSCapitalGainTaxReport() {
  Logger.log(JSON.stringify(getUSCapitalGainTaxReport(new Date().getFullYear()), null, 2));
}


// ============================================================
// webapp_04_dividend.gs — หน้าปันผล (รายงานวิเคราะห์ปันผลรายปี)
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
// ============================================================

// ──────────────────────────────────────
// รายงานปันผลรายปี แยก US/TH — เรียกจากหน้าเว็บผ่าน
// google.script.run.getDividendAnnualReport(year)
// ──────────────────────────────────────
function getDividendAnnualReport(year) {
  const yr = year || new Date().getFullYear();
  try {
    const sheet = getSheet(SHEETS.DIV);
    const lastRow = sheet ? sheet.getLastRow() : 0;

    if (!sheet || lastRow < START_ROW.DIV) {
      return { success: true, year: yr, us: [], th: [], usTotal: 0, thTotal: 0,
               usPortTotal: 0, thPortTotal: 0, usCostTotal: 0, thCostTotal: 0,
               usCostTotalNative: 0, exchangeRate: 0 };
    }

    const numRows = lastRow - START_ROW.DIV + 1;
    const rows = sheet.getRange(START_ROW.DIV, 1, numRows, 17).getValues();

    const usMap = {}, thMap = {};
    let usTotal = 0, thTotal = 0;

    rows.forEach(row => {
      const payDate = row[DIV_COL.PAY_DATE - 1];
      const ticker  = row[DIV_COL.TICKER  - 1];
      if (!ticker || !(payDate instanceof Date)) return;
      if (payDate.getFullYear() !== yr) return;

      const marketRaw = String(row[DIV_COL.MARKET - 1] || '').toUpperCase();
      const netTHB = Number(row[DIV_COL.NET_THB - 1]) || 0;
      const key = String(ticker).trim().toUpperCase();
      const isUS = marketRaw.indexOf('US') !== -1 || marketRaw.indexOf('สหรัฐ') !== -1 ||
                   marketRaw.indexOf('USD') !== -1;

      if (isUS) { usMap[key] = (usMap[key] || 0) + netTHB; usTotal += netTHB; }
      else      { thMap[key] = (thMap[key] || 0) + netTHB; thTotal += netTHB; }
    });

    // ── แปลง totalCost ของหุ้น US จาก USD → THB ก่อนใช้คำนวณ Yield on Cost ──
    // (Holdings เก็บ totalCost เป็นสกุลเงินต้นทาง แต่ divAmt เป็น THB เสมอ ต้องแปลงให้ตรงกันก่อนหาร)
    const exchangeRate = getFxRate();

    const usHoldings = getHoldings(SHEETS.US_HOLD);
    const thHoldings = getHoldings(SHEETS.TH_HOLD);
    const usPortTotal = usHoldings.reduce((s, h) => s + h.valueNow, 0);
    const thPortTotal = thHoldings.reduce((s, h) => s + h.valueNow, 0);
    const usCostTotalNative = usHoldings.reduce((s, h) => s + h.totalCost, 0);
    const usCostTotal = usCostTotalNative * exchangeRate; // ← THB แล้ว
    const thCostTotal = thHoldings.reduce((s, h) => s + h.totalCost, 0); // THB อยู่แล้ว ไม่ต้องคูณ

    const usPortMap = {}, usCostMap = {}; // usCostMap เก็บ { native, thb }
    usHoldings.forEach(h => {
      const k = String(h.ticker).toUpperCase();
      usPortMap[k] = h.valueNow;
      usCostMap[k] = { native: h.totalCost, thb: h.totalCost * exchangeRate };
    });
    const thPortMap = {}, thCostMap = {};
    thHoldings.forEach(h => {
      const k = String(h.ticker).toUpperCase();
      thPortMap[k] = h.valueNow;
      thCostMap[k] = { native: h.totalCost, thb: h.totalCost }; // TH: native = thb เท่ากันอยู่แล้ว
    });

    function buildList(map, total, portMap, portTotal, costMap) {
      return Object.keys(map).map(t => {
        const divAmt = map[t];
        const divPct = total > 0 ? (divAmt / total) * 100 : 0;
        const holdVal = portMap[t] || 0;
        const holdPct = portTotal > 0 ? (holdVal / portTotal) * 100 : 0;
        const costObj = costMap[t] || { native: 0, thb: 0 };
        const yieldOnCost = costObj.thb > 0 ? (divAmt / costObj.thb) * 100 : 0; // ← หารด้วย THB เสมอ ถูกต้องแล้ว
        return {
          ticker: t,
          divAmt: divAmt,
          divPct: divPct,
          holdVal: holdVal,
          holdPct: holdPct,
          diffPct: divPct - holdPct,
          inPortfolio: holdVal > 0,
          costVal: costObj.thb,          // THB — ใช้คำนวณ/legacy
          costValNative: costObj.native, // สกุลเงินต้นทาง — ใช้แสดงผล
          yieldOnCost: yieldOnCost
        };
      }).sort((a, b) => b.divAmt - a.divAmt);
    }

    return {
      success: true,
      year: yr,
      us: buildList(usMap, usTotal, usPortMap, usPortTotal, usCostMap),
      th: buildList(thMap, thTotal, thPortMap, thPortTotal, thCostMap),
      usTotal: usTotal,
      thTotal: thTotal,
      usPortTotal: usPortTotal,
      thPortTotal: thPortTotal,
      usCostTotal: usCostTotal,             // THB
      thCostTotal: thCostTotal,             // THB
      usCostTotalNative: usCostTotalNative, // USD ดิบ (เผื่อ frontend อยากโชว์รวมด้วย)
      exchangeRate: exchangeRate
    };
  } catch (e) {
    logError('getDividendAnnualReport', e);
    return { success: false, error: e.message, year: yr, us: [], th: [],
             usTotal: 0, thTotal: 0, usPortTotal: 0, thPortTotal: 0,
             usCostTotal: 0, thCostTotal: 0, usCostTotalNative: 0, exchangeRate: 0 };
  }
}

// ──────────────────────────────────────
// แนวโน้มปันผลรายเดือน เทียบปีก่อนหน้า (YoY) — แยก US/TH
// เรียกจากหน้าเว็บผ่าน google.script.run.getDividendMonthlyTrend(year)
// ──────────────────────────────────────
function getDividendMonthlyTrend(year) {
  const yr = year || new Date().getFullYear();
  const prevYr = yr - 1;
  const emptyBucket = () => ({ cur: Array(12).fill(0), prev: Array(12).fill(0) });

  try {
    const sheet = getSheet(SHEETS.DIV);
    const lastRow = sheet ? sheet.getLastRow() : 0;

    if (!sheet || lastRow < START_ROW.DIV) {
      return { success: true, year: yr, prevYear: prevYr, us: emptyBucket(), th: emptyBucket() };
    }

    const numRows = lastRow - START_ROW.DIV + 1;
    const rows = sheet.getRange(START_ROW.DIV, 1, numRows, 17).getValues();

    const us = emptyBucket(), th = emptyBucket();

    rows.forEach(row => {
      const payDate = row[DIV_COL.PAY_DATE - 1];
      const ticker  = row[DIV_COL.TICKER  - 1];
      if (!ticker || !(payDate instanceof Date)) return;

      const y = payDate.getFullYear();
      if (y !== yr && y !== prevYr) return;

      const m = payDate.getMonth(); // 0-11
      const marketRaw = String(row[DIV_COL.MARKET - 1] || '').toUpperCase();
      const netTHB = Number(row[DIV_COL.NET_THB - 1]) || 0;
      const isUS = marketRaw.indexOf('US') !== -1 || marketRaw.indexOf('สหรัฐ') !== -1 ||
                   marketRaw.indexOf('USD') !== -1;
      const bucket = isUS ? us : th;

      if (y === yr) bucket.cur[m] += netTHB;
      else          bucket.prev[m] += netTHB;
    });

    return { success: true, year: yr, prevYear: prevYr, us: us, th: th };
  } catch (e) {
    logError('getDividendMonthlyTrend', e);
    return { success: false, error: e.message, year: yr, prevYear: prevYr,
             us: emptyBucket(), th: emptyBucket() };
  }
}

// ──────────────────────────────────────
// รายชื่อปีที่มีข้อมูลปันผล — ใช้ทำตัวเลือกปีในหน้าเว็บ
// ──────────────────────────────────────
function getDividendYears() {
  const nowYr = new Date().getFullYear();
  try {
    const sheet = getSheet(SHEETS.DIV);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (!sheet || lastRow < START_ROW.DIV) return [nowYr];

    const numRows = lastRow - START_ROW.DIV + 1;
    const payDates = sheet.getRange(START_ROW.DIV, DIV_COL.PAY_DATE, numRows, 1).getValues();

    const years = new Set([nowYr]);
    payDates.forEach(r => { if (r[0] instanceof Date) years.add(r[0].getFullYear()); });

    return Array.from(years).sort((a, b) => b - a);
  } catch (e) {
    logError('getDividendYears', e);
    return [nowYr];
  }
}

/** ประวัติปันผลของหุ้นตัวเดียว แยกตามปี — ใช้ตอนแตะการ์ดในหน้ารายงานปันผล */
function getDividendHistoryForTicker(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const sheet = getSheet(SHEETS.DIV);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (!sheet || lastRow < START_ROW.DIV) {
      return { success: true, ticker, market, years: [], totalNetTHB: 0, totalPayments: 0 };
    }

    const numRows = lastRow - START_ROW.DIV + 1;
    const rows = sheet.getRange(START_ROW.DIV, 1, numRows, 17).getValues();

    const isTargetUS = market === 'US';
    const yearMap = {}; // { year: { totalNetTHB, totalGrossNative, payments: [] } }
    let totalNetTHB = 0, totalPayments = 0;

    rows.forEach(row => {
      const rowTicker = String(row[DIV_COL.TICKER - 1] || '').trim().toUpperCase();
      const payDate = row[DIV_COL.PAY_DATE - 1];
      if (rowTicker !== ticker || !(payDate instanceof Date)) return;

      const marketRaw = String(row[DIV_COL.MARKET - 1] || '').toUpperCase();
      const rowIsUS = marketRaw.indexOf('US') !== -1 || marketRaw.indexOf('สหรัฐ') !== -1 || marketRaw.indexOf('USD') !== -1;
      if (rowIsUS !== isTargetUS) return; // กันชื่อหุ้นซ้ำข้ามตลาด

      const yr = payDate.getFullYear();
      if (!yearMap[yr]) yearMap[yr] = { year: yr, totalNetTHB: 0, totalGrossNative: 0, payments: [] };

      const netThb = Number(row[DIV_COL.NET_THB - 1]) || 0;
      const grossNative = Number(row[DIV_COL.AMT - 1]) || 0;
      const perShare = Number(row[DIV_COL.PER_SHARE - 1]) || 0;
      const shares = Number(row[DIV_COL.SHARES - 1]) || 0;
      const currency = String(row[DIV_COL.CURRENCY - 1] || '');
      const round = String(row[DIV_COL.ROUND - 1] || '');
      const xdDate = row[DIV_COL.XD_DATE - 1];

      yearMap[yr].totalNetTHB += netThb;
      yearMap[yr].totalGrossNative += grossNative;
      yearMap[yr].payments.push({
        payDate: Utilities.formatDate(payDate, 'Asia/Bangkok', 'dd/MM/yyyy'),
        xdDate: (xdDate instanceof Date) ? Utilities.formatDate(xdDate, 'Asia/Bangkok', 'dd/MM/yyyy') : '',
        round, perShare, shares, grossNative, currency, netThb
      });

      totalNetTHB += netThb;
      totalPayments++;
    });

    const years = Object.values(yearMap)
      .sort((a, b) => b.year - a.year)
      .map(y => ({ ...y, payments: y.payments.sort((a, b) => new Date(b.payDate.split('/').reverse().join('-')) - new Date(a.payDate.split('/').reverse().join('-'))) }));

    // YoY growth เทียบปีก่อนหน้า
    years.forEach((y, i) => {
      const prev = years[i + 1]; // เรียงใหม่ล่าสุดก่อน ตัวถัดไปคือปีก่อนหน้า
      y.yoyPct = prev && prev.totalNetTHB > 0 ? ((y.totalNetTHB - prev.totalNetTHB) / prev.totalNetTHB) * 100 : null;
    });

    return { success: true, ticker, market, years, totalNetTHB, totalPayments,
      avgPerYear: years.length > 0 ? totalNetTHB / years.length : 0 };
  } catch (e) {
    logError('getDividendHistoryForTicker', e);
    return { success: false, error: e.message };
  }
}


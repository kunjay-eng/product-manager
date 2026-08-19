// ============================================================
// webapp_14_fee_analysis.gs — เมนู 💸 Fee Analysis (submenu ใต้ วิเคราะห์)
// ============================================================
// ดึงข้อมูลดิบจาก Transaction log ทั้ง US/TH โดยตรง (ไม่พึ่ง Realized P&L
// เพราะต้องนับ fee ของ "ทุกรายการ" ทั้งซื้อและขาย ไม่ใช่แค่รอบที่ขายแล้ว)
// แปลง USD→THB ด้วยอัตราแลกเปลี่ยนที่กรอกไว้ในแต่ละแถว (col Q ตามที่
// saveUSStock() เขียนไว้อยู่แล้ว)
//
// คืนค่า 3 ชุดในก้อนเดียว: all / US / TH — front-end สลับ toggle ได้โดย
// ไม่ต้องยิง request ใหม่ (เหมือน pattern หน้า Summary เดิมที่คืน d.th/d.us
// มาให้ในก้อนเดียวกัน)
// ============================================================

function getFeeAnalysisData() {
  try {
    const usRows = _readFeeRows(SHEETS.US_TRANS, 'US');
    const thRows = _readFeeRows(SHEETS.TH_TRANS, 'TH');
    const allRows = usRows.concat(thRows);

    return {
      success: true,
      updatedAt: new Date().toLocaleString('th-TH'),
      all: _computeFeeSummary(allRows, _getRealizedProfitByMarket('all')),
      US:  _computeFeeSummary(usRows,  _getRealizedProfitByMarket('US')),
      TH:  _computeFeeSummary(thRows,  _getRealizedProfitByMarket('TH'))
    };
  } catch (e) {
    logError('getFeeAnalysisData', e);
    return { success: false, error: e.message };
  }
}

/**
 * คำนวณ summary/monthly/byStock/feeVsProfit/overTrading จากชุดแถวที่ให้มา
 * (ใช้ซ้ำได้ทั้ง all/US/TH — ตรรกะเดียวกันทุกครั้ง แค่ input rows ต่างกัน)
 */
function _computeFeeSummary(rows, realizedProfitTHB) {
  if (!rows.length) {
    return {
      summary: { totalFeeTHB:0, totalCommissionTHB:0, totalOtherTHB:0, totalVatTHB:0, buyCount:0, sellCount:0, avgFeePerOrder:0 },
      monthly: [], byStock: [], feeVsProfit: null, overTrading: [], mostExpensiveStock: null, highestFrequencyStock: null
    };
  }

  let totalFee = 0, totalCommission = 0, totalOther = 0, totalVat = 0, buyCount = 0, sellCount = 0;
  rows.forEach(r => {
    totalFee += r.feeTHB;
    totalCommission += r.commissionTHB;
    totalOther += r.otherTHB;
    totalVat += r.vatTHB;
    if (r.type === 'ซื้อ') buyCount++; else if (r.type === 'ขาย') sellCount++;
  });

  // ── แยกตามเดือน ──
  const monthlyMap = {};
  rows.forEach(r => {
    const key = r.dateStr.substring(0, 7); // yyyy-MM
    monthlyMap[key] = (monthlyMap[key] || 0) + r.feeTHB;
  });
  const monthly = Object.keys(monthlyMap).sort().map(k => ({
    label: k, feeTHB: Math.round(monthlyMap[k] * 100) / 100
  }));

  // ── แยกตามหุ้น (รวม _c1, _c2, ... เข้ากับชื่อหุ้นเดิม — ถือเป็นหุ้นตัวเดียวกัน) ──
  const stockMap = {};
  rows.forEach(r => {
    const baseTicker = _stripCycleSuffix(r.ticker);
    const key = r.market + ':' + baseTicker;
    if (!stockMap[key]) stockMap[key] = { ticker: baseTicker, market: r.market, orders: 0, feeTHB: 0, shares: 0 };
    stockMap[key].orders++;
    stockMap[key].feeTHB += r.feeTHB;
    stockMap[key].shares += r.shares;
  });
  const byStock = Object.values(stockMap).map(s => ({
    ticker: s.ticker, market: s.market, orders: s.orders,
    feeTHB: Math.round(s.feeTHB * 100) / 100,
    feePerShare: s.shares > 0 ? Math.round((s.feeTHB / s.shares) * 10000) / 10000 : 0,
    feePerOrder: Math.round((s.feeTHB / s.orders) * 100) / 100
  })).sort((a, b) => b.feeTHB - a.feeTHB);

  const feeVsProfit = {
    realizedProfitTHB: realizedProfitTHB,
    totalFeeTHB: Math.round(totalFee * 100) / 100,
    feePct: realizedProfitTHB > 0 ? Math.round((totalFee / realizedProfitTHB) * 10000) / 100 : null
  };

  // ── Over Trading — orders เกิน threshold ──
  const ORDER_THRESHOLD = 15;
  const overTrading = byStock
    .filter(s => s.orders >= ORDER_THRESHOLD)
    .map(s => Object.assign({}, s, { status: s.orders >= ORDER_THRESHOLD * 2 ? 'high' : 'warn' }));

  return {
    summary: {
      totalFeeTHB: Math.round(totalFee * 100) / 100,
      totalCommissionTHB: Math.round(totalCommission * 100) / 100,
      totalOtherTHB: Math.round(totalOther * 100) / 100,
      totalVatTHB: Math.round(totalVat * 100) / 100,
      buyCount: buyCount, sellCount: sellCount,
      avgFeePerOrder: (buyCount + sellCount) > 0 ? Math.round((totalFee / (buyCount + sellCount)) * 100) / 100 : 0
    },
    monthly: monthly,
    byStock: byStock,
    feeVsProfit: feeVsProfit,
    overTrading: overTrading,
    mostExpensiveStock: byStock[0] || null,
    highestFrequencyStock: [...byStock].sort((a, b) => b.orders - a.orders)[0] || null
  };
}

/**
 * ตัด _c1, _c2, ... ท้ายชื่อหุ้นออก — เช่น "NVDA_c2" → "NVDA"
 * ใช้ตอนรวม fee ของหุ้นตัวเดียวกันที่เคยขายหมดแล้วซื้อใหม่หลายรอบ (ระบบ Cycle
 * ในเฟส 1) ให้นับเป็นหุ้นตัวเดียวกัน ไม่แยกเป็นหลายแถวใน Fee Analysis
 */
function _stripCycleSuffix(ticker) {
  return String(ticker || '').replace(/_C\d+$/i, '').trim();
}

/**
 * อ่านข้อมูลดิบจาก Transaction log 1 ตลาด แล้วคำนวณ fee THB ต่อแถว
 * US: commission(H)+taf(I)+tafFee(J)+vat(K) เป็น USD → คูณ exchangeRate(Q) เป็น THB
 * TH: commission(H)+fee(I)+vat(J) เป็น THB อยู่แล้ว
 */
function _readFeeRows(sheetName, market) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const startRow = START_ROW.HOLD; // แถวเริ่ม transaction log เดียวกับที่ saveUSStock/saveTHStock ใช้
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const isUS = market === 'US';
  const numCols = isUS ? 17 : 11;
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();

  const rows = [];
  data.forEach(r => {
    const ticker = String(r[2] || '').trim(); // col C
    if (!ticker) return;
    const dateVal = r[1]; // col B
    const type = String(r[3] || '').trim(); // col D
    const shares = parseFloat(r[4]) || 0; // col E

    let commission, other, vat, exchangeRate;
    if (isUS) {
      commission = parseFloat(r[7]) || 0; // col H
      const taf = parseFloat(r[8]) || 0;  // col I
      const tafFee = parseFloat(r[9]) || 0; // col J
      other = taf + tafFee;
      vat = parseFloat(r[10]) || 0; // col K
      exchangeRate = parseFloat(r[16]) || 0; // col Q
      if (!exchangeRate) return; // ไม่มีอัตราแลกเปลี่ยน แปลง THB ไม่ได้ ข้ามแถวนี้
    } else {
      commission = parseFloat(r[7]) || 0; // col H
      other = parseFloat(r[8]) || 0;      // col I
      vat = parseFloat(r[9]) || 0;        // col J
      exchangeRate = 1;
    }

    let dateStr;
    if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    else dateStr = String(dateVal || '').substring(0, 10);
    if (!dateStr) return;

 rows.push({
  ticker: ticker, market: market, type: type, shares: shares, dateStr: dateStr,
  price: parseFloat(r[5]) || 0,                       // col F ราคาต่อหุ้น
  commissionTHB: commission * exchangeRate,
  otherTHB: other * exchangeRate,
  vatTHB: vat * exchangeRate,
  feeTHB: (commission + other + vat) * exchangeRate,
  feeNative: commission + other + vat,                 // USD สำหรับ US, THB สำหรับ TH
  tradeValueNative: (parseFloat(r[5]) || 0) * shares    // มูลค่ารายการ (สกุลเงินต้นทาง)
});


  });
  return rows;
}

/**
 * กำไรที่รับรู้แล้ว (THB) แยกตาม market — ใช้ cell สรุปที่มีสูตรอยู่แล้วใน
 * ชีต Realized P&L: TH ใช้ TH_NET_THB (เป็น THB อยู่แล้ว), US ใช้ US_NET_USD
 * คูณ FX_RATE ที่ใช้คำนวณ ณ ตอนนั้น ให้เป็น THB, all ใช้ TOTAL_THB (รวมแล้ว)
 */
function _getRealizedProfitByMarket(market) {
  try {
    if (market === 'TH') {
      const sheet = getSheet(SHEETS.TH_REAL);
      if (!sheet) return 0;
      return parseFloat(sheet.getRange(REALIZED_CELL.TH_NET_THB).getValue()) || 0;
    }
    if (market === 'US') {
      const sheet = getSheet(SHEETS.US_REAL);
      if (!sheet) return 0;
      const usdProfit = parseFloat(sheet.getRange(REALIZED_CELL.US_NET_USD).getValue()) || 0;
      const fxRate = parseFloat(sheet.getRange(REALIZED_CELL.FX_RATE).getValue()) || 0;
      return usdProfit * fxRate;
    }
    const sheet = getSheet(SHEETS.US_REAL) || getSheet(SHEETS.TH_REAL);
    if (!sheet) return 0;
    return parseFloat(sheet.getRange(REALIZED_CELL.TOTAL_THB).getValue()) || 0;
  } catch (e) {
    return 0;
  }
}


/** สรุป fee ของหุ้นตัวเดียว (รวมทุกรอบ _C1, _C2...) — ใช้ในการ์ด Fee หน้า Stock Detail */
function getStockFeeCard(ticker, market) {
  try {
    const baseTicker = String(ticker || '').trim().toUpperCase();
    const sheetName = market === 'TH' ? SHEETS.TH_TRANS : SHEETS.US_TRANS;
    const allRows = _readFeeRows(sheetName, market); // ทั้งตลาด — ใช้ fallback sell-fee rate ด้วย
    const rows = allRows.filter(r => _stripCycleSuffix(r.ticker).toUpperCase() === baseTicker);

    if (!rows.length) {
      return { success: true, hasData: false, cycles: 0,
        currentCycleBuyFeeTHB: 0, totalBuyFeeTHB: 0, totalSellFeeTHB: 0, totalFeeTHB: 0,
        currentCycleBuyFeeNative: 0, totalBuyFeeNative: 0, totalSellFeeNative: 0, totalFeeNative: 0,
        avgFeePerOrderTHB: 0, avgFeePerCycleTHB: 0, avgFeePerOrderNative: 0, avgFeePerCycleNative: 0,
        buyCount: 0, sellCount: 0, overTrading: null, totalBuyShares: 0 };
    }

    const cycleOrder = r => { const m = String(r.ticker).match(/_C(\d+)$/i); return m ? parseInt(m[1],10) : 0; };
    const cycles = [...new Set(rows.map(r => r.ticker.toUpperCase()))];
    const latestCycleTicker = rows.reduce((a,b) => cycleOrder(b) > cycleOrder(a) ? b : a, rows[0]).ticker.toUpperCase();

    let totalBuy=0, totalSell=0, buyCount=0, sellCount=0, currentCycleBuy=0;
    let totalBuyN=0, totalSellN=0, currentCycleBuyN=0, totalBuyShares=0;
    rows.forEach(r => {
      if (r.type === 'ซื้อ') {
        totalBuy += r.feeTHB; totalBuyN += r.feeNative; totalBuyShares += r.shares; buyCount++;
        if (r.ticker.toUpperCase() === latestCycleTicker) { currentCycleBuy += r.feeTHB; currentCycleBuyN += r.feeNative; }
      } else if (r.type === 'ขาย') {
        totalSell += r.feeTHB; totalSellN += r.feeNative; sellCount++;
      }
    });

    const totalFee = totalBuy + totalSell;
    const totalFeeN = totalBuyN + totalSellN;
    const totalOrders = buyCount + sellCount;

    const ORDER_THRESHOLD = 15;
    const overTrading = totalOrders >= ORDER_THRESHOLD ? (totalOrders >= ORDER_THRESHOLD*2 ? 'high' : 'warn') : null;

    const sellFeeEst = _estimateSellFeeRate(rows, allRows);

    return {
      success: true, hasData: true, cycles: cycles.length,
      currentCycleBuyFeeTHB: Math.round(currentCycleBuy*100)/100,
      totalBuyFeeTHB: Math.round(totalBuy*100)/100,
      totalSellFeeTHB: Math.round(totalSell*100)/100,
      totalFeeTHB: Math.round(totalFee*100)/100,
      currentCycleBuyFeeNative: Math.round(currentCycleBuyN*100)/100,
      totalBuyFeeNative: Math.round(totalBuyN*100)/100,
      totalSellFeeNative: Math.round(totalSellN*100)/100,
      totalFeeNative: Math.round(totalFeeN*100)/100,
      avgFeePerOrderTHB: totalOrders>0 ? Math.round((totalFee/totalOrders)*100)/100 : 0,
      avgFeePerCycleTHB: cycles.length>0 ? Math.round((totalFee/cycles.length)*100)/100 : 0,
      avgFeePerOrderNative: totalOrders > 0 ? Math.round((totalFeeN / totalOrders) * 100) / 100 : 0,
      avgFeePerCycleNative: cycles.length > 0 ? Math.round((totalFeeN / cycles.length) * 100) / 100 : 0,
      buyCount, sellCount, overTrading,
      totalBuyShares: Math.round(totalBuyShares * 10000) / 10000,
      sellFeeRatePct: sellFeeEst ? Math.round(sellFeeEst.rate * 10000) / 100 : null,
      sellFeeRateSource: sellFeeEst ? sellFeeEst.source : null
    };
  } catch (e) {
    logError('getStockFeeCard', e);
    return { success: false, error: e.message };
  }
}


/** เติม breakeven price + fee vs profit % — เรียกหลัง getStockFeeCard() เมื่อรู้ต้นทุน/กำไรแล้ว */
function _attachFeeInsights(feeCard, avgCostNative, sharesRemain, unrealizedProfitNative, realizedProfitNative) {
  if (!feeCard || !feeCard.success || !feeCard.hasData) return feeCard;

  const costBasis = (avgCostNative * (sharesRemain || 0)) + feeCard.totalBuyFeeNative;

  if (sharesRemain > 0) {
    if (feeCard.sellFeeRatePct !== null) {
      const r = feeCard.sellFeeRatePct / 100;
      feeCard.breakevenPrice = r < 1 ? Math.round((costBasis / (sharesRemain * (1 - r))) * 100) / 100 : null;
      feeCard.breakevenIncludesSellFee = true;
    } else {
      feeCard.breakevenPrice = Math.round((costBasis / sharesRemain) * 100) / 100;
      feeCard.breakevenIncludesSellFee = false;
    }
  } else {
    feeCard.breakevenPrice = null;
  }

  // แบ่ง fee ซื้อตามสัดส่วนหุ้นที่ขายไปแล้ว vs ที่ยังถืออยู่
  const buyFeePerShare = feeCard.totalBuyShares > 0 ? feeCard.totalBuyFeeNative / feeCard.totalBuyShares : 0;
  const unrealizedFeeNative = buyFeePerShare * (sharesRemain || 0);
  const realizedFeeNative = feeCard.totalFeeNative - unrealizedFeeNative;

  feeCard.unrealizedFeeNative = Math.round(unrealizedFeeNative * 100) / 100;
  feeCard.realizedFeeNative = Math.round(realizedFeeNative * 100) / 100;

  feeCard.feeVsUnrealizedProfitPct = (unrealizedProfitNative !== null && unrealizedProfitNative > 0)
    ? Math.round((unrealizedFeeNative / unrealizedProfitNative) * 10000) / 100 : null;

  feeCard.feeVsRealizedProfitPct = (realizedProfitNative !== null && realizedProfitNative > 0)
    ? Math.round((realizedFeeNative / realizedProfitNative) * 10000) / 100 : null;

  return feeCard;
}



/** ประมาณ % fee ต่อมูลค่าขาย — ใช้ประวัติขายจริงของหุ้นตัวนี้ก่อน ถ้าไม่เคยขายเลยค่อย fallback เป็นค่าเฉลี่ยทั้งตลาด */
function _estimateSellFeeRate(stockRows, marketRows) {
  const ownSells = stockRows.filter(r => r.type === 'ขาย' && r.tradeValueNative > 0);
  if (ownSells.length) {
    const fee = ownSells.reduce((s,r) => s + r.feeNative, 0);
    const val = ownSells.reduce((s,r) => s + r.tradeValueNative, 0);
    return val > 0 ? { rate: fee/val, source: 'หุ้นตัวนี้เอง (' + ownSells.length + ' ครั้ง)' } : null;
  }
  const marketSells = marketRows.filter(r => r.type === 'ขาย' && r.tradeValueNative > 0);
  if (marketSells.length) {
    const fee = marketSells.reduce((s,r) => s + r.feeNative, 0);
    const val = marketSells.reduce((s,r) => s + r.tradeValueNative, 0);
    return val > 0 ? { rate: fee/val, source: 'ค่าเฉลี่ยทั้งตลาด (' + marketSells.length + ' ครั้ง)' } : null;
  }
  return null;
}


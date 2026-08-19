// ══════════════════════════════════════════════════════════════════
// webapp_18_trade_analysis.gs
// ประวัติการซื้อ-ขาย + วิเคราะห์การเทรด (History / Detail / Analysis)
// วางไฟล์นี้เป็นไฟล์ .gs ใหม่ในโปรเจกต์ Apps Script (ไม่ต้องแก้ไฟล์อื่น)
// ══════════════════════════════════════════════════════════════════


// ── ดึงธุรกรรมดิบจาก Transaction Log (US + TH) พร้อม tag รอบ (_cX) ──
// filters: { ticker, market('US'/'TH'), type('ซื้อ'/'ขาย'), startDate, endDate }
// คอลัมน์อ้างอิงตาม saveUSStock/saveTHStock: C=ticker D=ประเภท E=จำนวน F=ราคา
// H=ค่าคอม I/J/K=ค่าธรรมเนียมต่างๆ (US มี Q=อัตราแลกเปลี่ยน ณ วันทำรายการ)
function _getRawTransactionItems(filters) {
  filters = filters || {};
  const sources = [];
  if (!filters.market || filters.market === 'US') sources.push({ name: SHEETS.US_TRANS, market: 'US', cur: '$' });
  if (!filters.market || filters.market === 'TH') sources.push({ name: SHEETS.TH_TRANS, market: 'TH', cur: '฿' });

  let items = [];
  sources.forEach(src => {
    const rows = _getCachedSheetRows(src.name);
    rows.forEach(r => {
      const rawTicker = r[2];
      const type = r[3];
      if (!rawTicker || (type !== 'ซื้อ' && type !== 'ขาย')) return;

      const rawTickerUp = String(rawTicker).trim().toUpperCase();
      const cycleMatch = rawTickerUp.match(/_C(\d+)$/i);
      const cleanTicker = rawTickerUp.replace(/_C\d+$/i, '');
      const dateVal = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (isNaN(dateVal.getTime())) return;

      const shares = parseFloat(r[4]) || 0;
      const price  = parseFloat(r[5]) || 0;
      const commission = parseFloat(r[7]) || 0;
      const feeSum = src.market === 'US'
        ? (parseFloat(r[8]) || 0) + (parseFloat(r[9]) || 0) + (parseFloat(r[10]) || 0)
        : (parseFloat(r[8]) || 0) + (parseFloat(r[9]) || 0);
      const totalFee = commission + feeSum;

      items.push({
        date: Utilities.formatDate(dateVal, 'Asia/Bangkok', 'yyyy-MM-dd'),
        ticker: cleanTicker, market: src.market, cur: src.cur,
        cycle: cycleMatch ? Number(cycleMatch[1]) : null,
        type, shares, price, amount: shares * price, fee: totalFee,
        exchangeRate: src.market === 'US' ? (parseFloat(r[16]) || 0) : null // col Q
      });
    });
  });

  if (filters.ticker) {
    const t = String(filters.ticker).trim().toUpperCase();
    items = items.filter(x => x.ticker === t);
  }
  if (filters.type) items = items.filter(x => x.type === filters.type);
  if (filters.startDate) items = items.filter(x => x.date >= filters.startDate);
  if (filters.endDate)   items = items.filter(x => x.date <= filters.endDate);

  return items;
}


// ── รวมเงินปันผลทั้งหมด (กรองเฉพาะ ticker และ/หรือ ตลาด) — ใช้ใน History Analysis ──
// useTHB=true: ใช้คอลัมน์ NET_THB (แม่นยำเมื่อรวมหลายตลาด/สกุลเงินเข้าด้วยกัน)
// หมายเหตุ: คอลัมน์ตลาดในชีตปันผลเป็น free text (US/สหรัฐ/USD หรือ TH/ไทย/THB)
// เลยเช็คแบบ fuzzy .indexOf() เหมือนที่ getDividendAnnualReport ใช้อยู่ ไม่ใช่ exact match
function _getTotalDividends(ticker, useTHB, market) {
  try {
    const rows = getSheet(SHEETS.DIV).getDataRange().getValues();
    let total = 0;
    rows.forEach(r => {
      const t = r[DIV_COL.TICKER - 1];
      const rowMarketRaw = String(r[DIV_COL.MARKET - 1] || '').toUpperCase();
      const amt = parseFloat(r[useTHB ? DIV_COL.NET_THB - 1 : DIV_COL.AMT - 1]);
      if (!t || isNaN(amt)) return;
      if (ticker && String(t).trim().toUpperCase() !== String(ticker).trim().toUpperCase()) return;

      if (market) {
        const isUSRow = rowMarketRaw.indexOf('US') !== -1 || rowMarketRaw.indexOf('สหรัฐ') !== -1;
        const wantUS  = String(market).trim().toUpperCase() === 'US';
        if (isUSRow !== wantUS) return;
      }

      total += amt;
    });
    return total;
  } catch (e) {
    logError('_getTotalDividends', e);
    return 0;
  }
}

// ── ทำให้ธุรกรรมทั้งหมดอยู่ในสกุลเงินเดียวกันก่อนรวมยอด ──
// ถ้ามีแค่ตลาดเดียว (กรอง ticker หรือบังเอิญมีแต่ US/TH) ใช้สกุลเดิมได้เลย
// ถ้ามีทั้ง US+TH ปนกัน แปลง USD → THB ด้วยอัตราแลกเปลี่ยนปัจจุบัน แล้วรายงานเป็น ฿ ทั้งหมด
function _normalizeRawForReporting(raw) {
  const markets = [...new Set(raw.map(x => x.market))];
  if (markets.length <= 1) {
    return { items: raw, cur: raw.length ? raw[0].cur : '$', mixed: false, fxRate: 1 };
  }
  const fxRate = getFxRate();
  const items = raw.map(x => x.market === 'US'
    ? Object.assign({}, x, { amount: x.amount * fxRate, fee: x.fee * fxRate, cur: '฿' })
    : Object.assign({}, x, { cur: '฿' }));
  return { items, cur: '฿', mixed: true, fxRate };
}


// ── ปันผลสะสมของ ticker เดียว (สกุลต้นทาง + THB) — ใช้ใน Ticker Detail ──
function _getDividendTotals(ticker) {
  try {
    const rows = getSheet(SHEETS.DIV).getDataRange().getValues();
    let native = 0, thb = 0;
    rows.forEach(r => {
      const t = r[DIV_COL.TICKER - 1];
      if (!t || String(t).trim().toUpperCase() !== String(ticker).trim().toUpperCase()) return;
      native += parseFloat(r[DIV_COL.AMT - 1]) || 0;
      thb    += parseFloat(r[DIV_COL.NET_THB - 1]) || 0;
    });
    return { native, thb };
  } catch (e) {
    logError('_getDividendTotals', e);
    return { native: 0, thb: 0 };
  }
}


// ── Snapshot ตำแหน่งปัจจุบันจากชีต Holdings (ราคาล่าสุด/ต้นทุนเฉลี่ย/กำไรยังไม่รับรู้) ──
function _getHoldingSnapshot(ticker, market) {
  const sheetName = market === 'US' ? SHEETS.US_HOLD : SHEETS.TH_HOLD; // แก้บั๊ก: เดิมชี้ไปชีต Realized P&L ผิด
  const rows = getHoldings(sheetName);
  const t = String(ticker).trim().toUpperCase();
  const row = rows.find(r => r.ticker.trim().toUpperCase() === t);
  if (!row) return null;
  return {
    priceNow: row.priceNow, avgCost: row.avgCost, sharesRemain: row.sharesRemain,
    unrealizedPL: row.unrealizedPL, unrealizedPct: row.unrealizedPct
  };
}


// ── จัดกลุ่มธุรกรรมเป็น "รอบการลงทุน" ต่อ ticker+market+_cX ──
// ใช้ avg-cost method คิดกำไร/ขาดทุนต่อการขายแต่ละครั้ง
function _buildTradeCycles(raw) {
  const groups = {};
  raw.forEach(x => {
    const key = x.market + '|' + x.ticker + '|' + (x.cycle || 'single');
    if (!groups[key]) groups[key] = { ticker: x.ticker, market: x.market, cur: x.cur, cycle: x.cycle, txns: [] };
    groups[key].txns.push(x);
  });

  return Object.values(groups).map(g => {
    g.txns.sort((a, b) => a.date.localeCompare(b.date));
    const buys  = g.txns.filter(t => t.type === 'ซื้อ');
    const sells = g.txns.filter(t => t.type === 'ขาย');

    const totalBuyShares  = buys.reduce((s, t) => s + t.shares, 0);
    const totalSellShares = sells.reduce((s, t) => s + t.shares, 0);
    const totalBuyAmount  = buys.reduce((s, t) => s + t.amount, 0);
    const totalSellAmount = sells.reduce((s, t) => s + t.amount, 0);
    const totalFee        = g.txns.reduce((s, t) => s + t.fee, 0);
    const capitalInvested = totalBuyAmount + buys.reduce((s, t) => s + t.fee, 0);

    const isCompleted = totalBuyShares > 0 && totalSellShares >= totalBuyShares - 0.0001;
    const entryDate = buys.length ? buys[0].date : (g.txns[0] ? g.txns[0].date : null);
    const exitDate  = isCompleted && sells.length ? sells[sells.length - 1].date : null;
    const holdingDays = (isCompleted && entryDate && exitDate)
      ? Math.round((new Date(exitDate) - new Date(entryDate)) / 86400000)
      : null;
    const realizedPL = isCompleted ? (totalSellAmount - totalBuyAmount - totalFee) : null;

    // รายละเอียดต่อการขาย: กำไร/ขาดทุนต่อไม้ (avg-cost) + บอกว่าขายบางส่วนหรือหมด
    let remainingShares = 0, cumBuyShares = 0, cumBuyCost = 0;
    const sellDetails = [];
    g.txns.forEach(t => {
      if (t.type === 'ซื้อ') {
        remainingShares += t.shares;
        cumBuyShares += t.shares;
        cumBuyCost += t.amount + t.fee;
      } else {
        const avgCost = cumBuyShares ? cumBuyCost / cumBuyShares : 0;
        const pl = (t.price - avgCost) * t.shares - t.fee;
        remainingShares -= t.shares;
        sellDetails.push({ date: t.date, shares: t.shares, price: t.price, avgCostAtSell: avgCost, pl, isFullExit: remainingShares <= 0.0001 });
      }
    });

    return {
      ticker: g.ticker, market: g.market, cur: g.cur, cycle: g.cycle,
      buys, sells, sellDetails,
      totalBuyShares, totalSellShares, totalBuyAmount, totalSellAmount, totalFee,
      capitalInvested, isCompleted, entryDate, exitDate, holdingDays, realizedPL
    };
  });
}


// ── Max Drawdown จากมูลค่าพอร์ตจริงรายวัน (Portfolio_Value_Log) ──
// ใช้ได้เฉพาะมุมมอง "ทั้งพอร์ต" เพราะ log เก็บมูลค่ารวมทั้งบัญชี ไม่แยกราย ticker
function _computePortfolioDrawdown() {
  try {
    const hist = getPortfolioValueHistory();
    if (!hist.success || !hist.points || hist.points.length < 2) {
      return { available: false, reason: hist.note || 'ข้อมูลยังไม่พอ (ต้องมีอย่างน้อย 2 วัน) — กดเริ่มเก็บข้อมูลก่อน' };
    }

    let peak = hist.points[0].value, peakDate = hist.points[0].date;
    let maxDD = 0, maxDDPct = 0, maxDDPeakDate = peakDate, maxDDTroughDate = peakDate;

    hist.points.forEach(pt => {
      if (pt.value > peak) { peak = pt.value; peakDate = pt.date; }
      const dd = peak - pt.value;
      const ddPct = peak > 0 ? dd / peak * 100 : 0;
      if (dd > maxDD) { maxDD = dd; maxDDPct = ddPct; maxDDPeakDate = peakDate; maxDDTroughDate = pt.date; }
    });

    const latest = hist.points[hist.points.length - 1];
    const currentDD = peak - latest.value;

    return {
      available: true,
      maxDrawdownTHB: maxDD,
      maxDrawdownPct: maxDDPct,
      peakDate: maxDDPeakDate,
      troughDate: maxDDTroughDate,
      currentDrawdownTHB: currentDD,
      currentDrawdownPct: peak > 0 ? currentDD / peak * 100 : 0,
      currentPeak: peak,
      latestValue: latest.value,
      dataPoints: hist.points.length
    };
  } catch (e) {
    logError('_computePortfolioDrawdown', e);
    return { available: false, reason: e.message };
  }
}


// ══════════════════════════════════════════════════════════════════
// ENTRY POINTS — เรียกจาก Index.html ผ่าน google.script.run
// ══════════════════════════════════════════════════════════════════

// ── หน้า List รวม (ทุก ticker พร้อม filter) ──
function getTransactionHistory(filters) {
  filters = filters || {};
  try {
    const items = _getRawTransactionItems(filters);
    items.sort((a, b) => b.date.localeCompare(a.date));

    // ── จัดกลุ่มตาม ticker+market — รวมทุกรอบ (_c1, _c2, ...) เข้าด้วยกัน ──
    const groupMap = {};
    items.forEach(x => {
      const key = x.market + '|' + x.ticker;
      if (!groupMap[key]) {
        groupMap[key] = { ticker: x.ticker, market: x.market, cur: x.cur, items: [], buyAmount: 0, sellAmount: 0, buyShares: 0, sellShares: 0, fee: 0 };
      }
      const g = groupMap[key];
      g.items.push(x);
      g.fee += x.fee;
      if (x.type === 'ซื้อ') { g.buyAmount += x.amount; g.buyShares += x.shares; }
      else                   { g.sellAmount += x.amount; g.sellShares += x.shares; }
    });

    const groups = Object.values(groupMap).map(g => {
      g.items.sort((a, b) => b.date.localeCompare(a.date));
      g.netAmount = g.sellAmount - g.buyAmount;
      g.avgBuyPrice  = g.buyShares  ? g.buyAmount  / g.buyShares  : 0;
      g.avgSellPrice = g.sellShares ? g.sellAmount / g.sellShares : 0;
      g.cycles = [...new Set(g.items.map(x => x.cycle).filter(c => c !== null))].sort((a, b) => a - b);
      return g;
    });
    groups.sort((a, b) => b.items[0].date.localeCompare(a.items[0].date));

    const buys = items.filter(x => x.type === 'ซื้อ');
    const sells = items.filter(x => x.type === 'ขาย');
    const summary = {
      totalBuyAmount: buys.reduce((s, x) => s + x.amount, 0),
      totalSellAmount: sells.reduce((s, x) => s + x.amount, 0),
      totalFee: items.reduce((s, x) => s + x.fee, 0),
      buyCount: buys.length,
      sellCount: sells.length,
      avgBuyPrice: buys.length ? buys.reduce((s, x) => s + x.price * x.shares, 0) / buys.reduce((s, x) => s + x.shares, 0) : 0
    };

    return { success: true, items, groups, summary };
  } catch (e) {
    logError('getTransactionHistory', e);
    return { success: false, error: e.message, items: [], groups: [] };
  }
}


// ── หน้ารายละเอียดต่อ ticker เดียว (ตัวหลักที่ใช้งานจริง) ──
function getTickerTransactionDetail(ticker, market) {
  try {
    ticker = String(ticker).trim().toUpperCase();
    const raw = _getRawTransactionItems({ ticker, market });
    if (!raw.length) return { success: false, error: 'ไม่พบธุรกรรมของ ' + ticker };

    const cur = raw[0].cur;
    const fxRate = market === 'US' ? getFxRate() : 1;
    const cycles = _buildTradeCycles(raw);
    const closedGroups = cycles.filter(c => c.cycle); // มีเลขรอบ = ปิดไปแล้ว

    const hold = _getHoldingSnapshot(ticker, market);
    const div = _getDividendTotals(ticker);

    const allBuys  = raw.filter(x => x.type === 'ซื้อ');
    const allSells = raw.filter(x => x.type === 'ขาย');
    const buyTotal  = allBuys.reduce((s, x) => s + x.amount, 0);
    const sellTotal = allSells.reduce((s, x) => s + x.amount, 0);
    const totalFee  = raw.reduce((s, x) => s + x.fee, 0);
    const realizedProfit = closedGroups.reduce((s, c) => s + (c.realizedPL || 0), 0);
    const overallAvgBuyPrice = hold ? hold.avgCost : (buyTotal / (allBuys.reduce((s, x) => s + x.shares, 0) || 1));

    const buyDates = allBuys.map(x => x.date).sort();
    const buyPrices = allBuys.map(x => x.price);
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

    // ── รายการต่อไม้ ทุกกลุ่มรวมกัน ──
    const transactions = [];
    cycles.forEach(g => {
      const cycleLabel = g.cycle ? ('รอบที่ ' + g.cycle) : 'ถือครอง';
      const isOpenGroup = !g.cycle;

      g.buys.forEach(t => {
        const diffPct = overallAvgBuyPrice ? (t.price - overallAvgBuyPrice) / overallAvgBuyPrice * 100 : 0;
        let plRow = null, plPctRow = null, holdingDaysRow = null;
        if (isOpenGroup && hold) {
          plRow = (hold.priceNow - t.price) * t.shares - t.fee;
          plPctRow = t.price ? (hold.priceNow - t.price) / t.price * 100 : 0;
          holdingDaysRow = Math.round((new Date(today) - new Date(t.date)) / 86400000);
        }
        transactions.push({
          date: t.date, type: 'ซื้อ', cycleLabel, isOpenGroup,
          price: t.price, shares: t.shares, amount: t.amount, fee: t.fee,
          netAmount: t.amount + t.fee, exchangeRate: t.exchangeRate || null,
          diffFromAvgPct: diffPct, pl: plRow, plPct: plPctRow, holdingDays: holdingDaysRow
        });
      });

      g.sellDetails.forEach(sd => {
        const plPct = sd.avgCostAtSell ? (sd.price - sd.avgCostAtSell) / sd.avgCostAtSell * 100 : 0;
        const holdingDaysRow = g.entryDate ? Math.round((new Date(sd.date) - new Date(g.entryDate)) / 86400000) : null;
        transactions.push({
          date: sd.date, type: 'ขาย', cycleLabel, isOpenGroup,
          price: sd.price, shares: sd.shares, amount: sd.price * sd.shares,
          fee: 0, netAmount: sd.price * sd.shares,
          exchangeRate: null, diffFromAvgPct: null,
          pl: sd.pl, plPct, holdingDays: holdingDaysRow, isFullExit: sd.isFullExit
        });
      });
    });
    transactions.sort((a, b) => b.date.localeCompare(a.date));

    return {
      success: true, ticker, market, cur, fxRate,
      summary: {
        buyTotal, buyTotalTHB: market === 'US' ? buyTotal * fxRate : buyTotal,
        sellTotal, sellTotalTHB: market === 'US' ? sellTotal * fxRate : sellTotal,
        realizedProfit, realizedProfitTHB: market === 'US' ? realizedProfit * fxRate : realizedProfit,
        unrealizedProfit: hold ? hold.unrealizedPL : 0,
        unrealizedProfitTHB: hold ? (market === 'US' ? hold.unrealizedPL * fxRate : hold.unrealizedPL) : 0,
        unrealizedPct: hold ? hold.unrealizedPct : 0,
        totalFee, totalFeeTHB: market === 'US' ? totalFee * fxRate : totalFee,
        totalDividendNative: div.native, totalDividendTHB: div.thb,
        sharesHeld: hold ? hold.sharesRemain : 0,
        buyCount: allBuys.length,
        firstBuyDate: buyDates[0] || null,
        lastBuyDate: buyDates[buyDates.length - 1] || null,
        minBuyPrice: buyPrices.length ? Math.min(...buyPrices) : 0,
        maxBuyPrice: buyPrices.length ? Math.max(...buyPrices) : 0,
        holdingDaysTotal: buyDates.length ? Math.round((new Date(today) - new Date(buyDates[0])) / 86400000) : 0
      },
      transactions
    };
  } catch (e) {
    logError('getTickerTransactionDetail', e);
    return { success: false, error: e.message };
  }
}


// ── หน้า History Analysis (สถิติ 5 กลุ่ม รวมทุกหุ้น หรือกรองเฉพาะ ticker) ──
function getTradeAnalysis(filters) {
  filters = filters || {};
  try {
    const raw0 = _getRawTransactionItems(filters);
    const norm = _normalizeRawForReporting(raw0); // { items, cur, mixed, fxRate }
    const raw = norm.items; // ใช้ตัวนี้คำนวณสถิติรวม (สกุลเงินเดียวกันทั้งหมด)

    const cycles = _buildTradeCycles(raw);
    // cyclesNative: สกุลเงินจริงต่อรอบ (ไม่แปลง) — ใช้โชว์การ์ดรอบดีที่สุด/แย่ที่สุด
    const cyclesNative = norm.mixed ? _buildTradeCycles(raw0) : cycles;

    const completed = cycles.filter(c => c.isCompleted);
    const sortedCompleted = [...completed].sort((a, b) => a.exitDate.localeCompare(b.exitDate));

    const wins   = completed.filter(c => c.realizedPL > 0);
    const losses = completed.filter(c => c.realizedPL < 0);
    const grossProfit = wins.reduce((s, c) => s + c.realizedPL, 0);
    const grossLoss   = Math.abs(losses.reduce((s, c) => s + c.realizedPL, 0));

    // ── 1. ประสิทธิภาพการซื้อขาย ──
    const performance = {
      buyCount: raw.filter(x => x.type === 'ซื้อ').length,
      sellCount: raw.filter(x => x.type === 'ขาย').length,
      completedTrades: completed.length,
      winCount: wins.length,
      lossCount: losses.length,
      winRate: completed.length ? wins.length / completed.length * 100 : 0,
      lossRate: completed.length ? losses.length / completed.length * 100 : 0,
      avgGain: wins.length ? grossProfit / wins.length : 0,
      avgLoss: losses.length ? grossLoss / losses.length : 0,
      profitFactor: grossLoss > 0 ? grossProfit / grossLoss : (grossProfit > 0 ? Infinity : 0),
      expectancy: completed.length
        ? (wins.length / completed.length) * (wins.length ? grossProfit / wins.length : 0)
          - (losses.length / completed.length) * (losses.length ? grossLoss / losses.length : 0)
        : 0
    };

    // ── 2. การถือครอง ──
    const holdingDaysArr = sortedCompleted.map(c => c.holdingDays).filter(d => d != null);
    const holding = {
      avgHoldingDays: holdingDaysArr.length ? holdingDaysArr.reduce((a, b) => a + b, 0) / holdingDaysArr.length : 0,
      minHoldingDays: holdingDaysArr.length ? Math.min(...holdingDaysArr) : 0,
      maxHoldingDays: holdingDaysArr.length ? Math.max(...holdingDaysArr) : 0,
      latestHoldingDays: sortedCompleted.length ? sortedCompleted[sortedCompleted.length - 1].holdingDays : null,
      avgCapitalPerCycle: cycles.length ? cycles.reduce((s, c) => s + c.capitalInvested, 0) / cycles.length : 0
    };

    // ── 3. พฤติกรรมการซื้อ ──
    let avgDownCount = 0, chaseUpCount = 0, singleBuyCycles = 0, multiBuyCycles = 0;
    cycles.forEach(c => {
      if (c.buys.length === 1) singleBuyCycles++; else if (c.buys.length > 1) multiBuyCycles++;
      for (let i = 1; i < c.buys.length; i++) {
        if (c.buys[i].price < c.buys[i - 1].price) avgDownCount++;
        else if (c.buys[i].price > c.buys[i - 1].price) chaseUpCount++;
      }
    });
    const totalBuys = raw.filter(x => x.type === 'ซื้อ');
    const buying = {
      avgDownCount, chaseUpCount, singleBuyCycles, multiBuyCycles,
      avgBuysPerCycle: cycles.length ? totalBuys.length / cycles.length : 0,
      avgBuyValue: totalBuys.length ? totalBuys.reduce((s, x) => s + x.amount, 0) / totalBuys.length : 0
    };

    // ── 4. พฤติกรรมการขาย ──
    let profitSells = 0, lossSells = 0, partialSells = 0, fullSells = 0;
    cycles.forEach(c => c.sellDetails.forEach(sd => {
      if (sd.pl > 0) profitSells++; else if (sd.pl < 0) lossSells++;
      if (sd.isFullExit) fullSells++; else partialSells++;
    }));
    const selling = {
      profitSells, lossSells, partialSells, fullSells,
      maxGainPerCycle: completed.length ? Math.max(...completed.map(c => c.realizedPL)) : 0,
      maxLossPerCycle: completed.length ? Math.min(...completed.map(c => c.realizedPL)) : 0
    };

    // ── 5. สถิติผลตอบแทน ──
    const totalFeesAll = raw.reduce((s, x) => s + x.fee, 0);
    const totalDividends = _getTotalDividends(filters.ticker, norm.cur === '฿', filters.market);
    const cumulativeRealized = completed.reduce((s, c) => s + c.realizedPL, 0);
    const totalCapitalDeployed = cycles.reduce((s, c) => s + c.capitalInvested, 0);
    const roi = totalCapitalDeployed ? cumulativeRealized / totalCapitalDeployed * 100 : 0;

    const allDates = raw.map(x => x.date).sort();
    let cagr = null;
    if (allDates.length >= 2 && totalCapitalDeployed > 0) {
      const years = (new Date(allDates[allDates.length - 1]) - new Date(allDates[0])) / 86400000 / 365;
      if (years >= 1) cagr = (Math.pow((totalCapitalDeployed + cumulativeRealized) / totalCapitalDeployed, 1 / years) - 1) * 100;
    }

    let running = 0, peak = 0, maxDD = 0;
    sortedCompleted.forEach(c => {
      running += c.realizedPL;
      if (running > peak) peak = running;
      maxDD = Math.max(maxDD, peak - running);
    });

    const monthly = {}, yearly = {};
    completed.forEach(c => {
      const ym = c.exitDate.slice(0, 7), y = c.exitDate.slice(0, 4);
      monthly[ym] = (monthly[ym] || 0) + c.realizedPL;
      yearly[y]   = (yearly[y] || 0) + c.realizedPL;
    });

    const returns = {
      cumulativeRealizedProfit: cumulativeRealized,
      totalFees: totalFeesAll,
      totalDividends,
      roi, cagr,
      maxDrawdown: maxDD, // อิงกำไรสะสมจากรอบที่ปิดแล้ว
      portfolioDrawdown: filters.ticker
        ? { available: false, reason: 'ดูได้เฉพาะมุมมองทั้งพอร์ต (ไม่กรอง ticker)' }
        : _computePortfolioDrawdown(), // อิงมูลค่าพอร์ตจริงรายวัน
      monthly: Object.entries(monthly).sort((a, b) => a[0].localeCompare(b[0])).map(([k, v]) => ({ period: k, pl: v })),
      yearly: Object.entries(yearly).sort((a, b) => a[0].localeCompare(b[0])).map(([k, v]) => ({ period: k, pl: v }))
    };

    // ── สรุปต่อรอบ (เรียงกำไรมาก→น้อย) — ใช้สกุลเงินจริงของแต่ละรอบ ──
    const completedNative = cyclesNative.filter(c => c.isCompleted);
    const cyclesSummary = completedNative.map(c => ({
      ticker: c.ticker, market: c.market, cur: c.cur, cycle: c.cycle,
      entryDate: c.entryDate, exitDate: c.exitDate, holdingDays: c.holdingDays,
      buyCount: c.buys.length, capitalInvested: c.capitalInvested,
      realizedPL: c.realizedPL,
      roi: c.capitalInvested ? c.realizedPL / c.capitalInvested * 100 : 0,
      profitPerDay: c.holdingDays ? c.realizedPL / c.holdingDays : c.realizedPL
    })).sort((a, b) => b.realizedPL - a.realizedPL);

    return {
      success: true,
      cur: norm.cur, mixedCurrency: norm.mixed, // ใช้ให้ frontend ต่อสัญลักษณ์เงิน $/฿
      performance, holding, buying, selling, returns,
      bestCycle: cyclesSummary[0] || null,
      worstCycle: cyclesSummary[cyclesSummary.length - 1] || null,
      cyclesSummary,
      cycleCount: cycles.length
    };
  } catch (e) {
    logError('getTradeAnalysis', e);
    return { success: false, error: e.message };
  }
}

// ── อ่านชีตผ่าน cache (ลดการอ่านซ้ำเวลาเปิดหลาย modal ต่อเนื่อง) ──
function _getCachedSheetRows(sheetName, ttlSeconds) {
  ttlSeconds = ttlSeconds || 180;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'rows_' + sheetName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 200);
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* เพิกเฉย → อ่านจากชีตจริงแทน */ }

  const rows = getSheet(sheetName).getDataRange().getValues();
  try {
    cache.put(cacheKey, JSON.stringify(rows), ttlSeconds);
  } catch (e) {
    // ข้อมูลใหญ่เกิน 100KB/key — ข้าม cache ได้ ระบบยังทำงานถูกต้อง แค่ไม่เร็วขึ้น
  }
  return rows;
}

// ── เคลียร์ cache — เรียกทันทีหลังบันทึกซื้อ/ขายใหม่ กัน user เห็นข้อมูลเก่าค้าง ──
function _clearTransactionCache() {
  CacheService.getScriptCache().removeAll([
    'rows_' + SHEETS.US_TRANS.replace(/[^a-zA-Z0-9]/g, '_'),
    'rows_' + SHEETS.TH_TRANS.replace(/[^a-zA-Z0-9]/g, '_')
  ]);
}


// ── ปุ่ม "เริ่มเก็บข้อมูลมูลค่าพอร์ต" — เรียกจากหน้า History Analysis ──
// (ใช้ initPortfolioHistoryTracking() ที่มีอยู่แล้วใน webapp_15_portfolio_history.gs)


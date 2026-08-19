// ══════════════════════════════════════════════════════════════════
// webapp_25_compare.gs
// เปรียบเทียบหุ้น (Compare) — เทียบได้สูงสุด 5 ตัว จาก Holdings + Watchlist
// รองรับ 2 มุมมอง: Holding Period (ตั้งแต่ซื้อจริง) และ Same Period (ช่วงเวลาเดียวกัน)
// วางไฟล์นี้เป็นไฟล์ .gs ใหม่ในโปรเจกต์ (ไม่ต้องแก้ไฟล์อื่น)
// ══════════════════════════════════════════════════════════════════

const CMP_BENCHMARK_SYMBOL = { US: 'SPY', TH: 'SET' };

// ── จุดเข้าหลัก เรียกจาก frontend ──
// stocks: [{ ticker, market('US'/'TH'), kind('holding'/'watch') }, ...] สูงสุด 5 ตัว
// period: 'holding' | '1M' | '3M' | '6M' | 'YTD' | '1Y'
function getCompareData(stocks, period) {
  try {
    if (!Array.isArray(stocks) || !stocks.length) return { success: false, error: 'กรุณาเลือกหุ้นอย่างน้อย 1 ตัว' };
    if (stocks.length > 5) stocks = stocks.slice(0, 5);

    const now = new Date();
    const benchmarkCache = {};
    const results = [];

    stocks.forEach(st => {
      const ticker = String(st.ticker || '').trim().toUpperCase();
      const market = st.market;
      const kind = st.kind === 'watch' ? 'watch' : 'holding';

      const raw = _cmpGetOHLCSeries(ticker, market, kind);
      if (!raw || !raw.length) {
        results.push({ ticker, market, kind, available: false, reason: 'ไม่มีข้อมูลราคาย้อนหลังของ ' + ticker });
        return;
      }

      const cur = market === 'US' ? '$' : '฿';
      const currentPrice = raw[raw.length - 1].close;

      // ── Holding Period (เฉพาะหุ้นที่ถือจริง — Watchlist ยังไม่ได้ซื้อ ไม่มีวันซื้อให้อ้างอิง) ──
      let holdingPeriod = { available: false, reason: kind === 'watch' ? 'ยังไม่ได้ซื้อ (อยู่ใน Watchlist)' : 'ไม่พบวันที่ซื้อ' };
      let investmentResult = null;
      let firstBuyDate = null;
      let detail = null; // เก็บไว้ใช้ซ้ำตอนสร้าง markers (Phase 2) กันดึงข้อมูลซ้ำ

      if (kind === 'holding') {
        detail = getTickerTransactionDetail(ticker, market);
        if (detail && detail.success && detail.summary.firstBuyDate) {
          firstBuyDate = new Date(detail.summary.firstBuyDate);
          const startIdx = _cmpFindStartIndex(raw, firstBuyDate);
          if (startIdx >= 0) {
            const startPrice = raw[startIdx].close;
            const returnPct = ((currentPrice - startPrice) / startPrice) * 100;
            const days = Math.max(1, Math.round((now - firstBuyDate) / 86400000));
            holdingPeriod = {
              available: true, startDate: detail.summary.firstBuyDate, days,
              returnPct, cagr: _cmpCAGR(returnPct, days)
            };
          }
          investmentResult = _cmpGetInvestmentResult(detail, market, currentPrice);
        }
      }

      // ── Same Period (ทุกหุ้น รวม Watchlist ด้วย) ──
      const spStart = _cmpResolvePeriodStart(period === 'holding' ? '3M' : period); // ถ้าโหมด holding ใช้ 3M เป็นค่าสำรองของ samePeriod
      const spIdx = _cmpFindStartIndex(raw, spStart);
      let samePeriod = { available: false };
      if (spIdx >= 0) {
        const startPrice = raw[spIdx].close;
        const returnPct = ((currentPrice - startPrice) / startPrice) * 100;
        const days = Math.max(1, Math.round((now - raw[spIdx].date) / 86400000));
        samePeriod = { available: true, startDate: _cmpFmtDate(raw[spIdx].date), days, returnPct, cagr: _cmpCAGR(returnPct, days) };
      }

      // ── เลือกช่วงข้อมูลสำหรับกราฟ ตามโหมดที่ผู้ใช้เลือก ──
      let chartStartIdx = -1;
      if (period === 'holding') {
        if (kind === 'holding' && holdingPeriod.available) chartStartIdx = _cmpFindStartIndex(raw, firstBuyDate);
      } else {
        chartStartIdx = spIdx;
      }

      const series = _cmpBuildSeries(raw, chartStartIdx);
      const drawdown = _cmpBuildDrawdown(series);

      // ── Phase 2: จุด Buy/Sell/Trailing Stop บนกราฟ (เฉพาะหุ้นที่ถือจริง และมีกราฟให้วางจุด) ──
      const markers = kind === 'holding' ? _cmpBuildMarkers(ticker, market, detail, series) : null;

      // ── Benchmark (SPY สำหรับหุ้น US, SET สำหรับหุ้นไทย) — rebase 100 จากวันเริ่มต้นเดียวกับหุ้นตัวนี้ ──
      let benchmark = null;
      if (chartStartIdx >= 0) {
        const symbol = CMP_BENCHMARK_SYMBOL[market];
        const bmRaw = symbol ? _cmpGetBenchmarkSeries(market, raw[chartStartIdx].date, now, benchmarkCache) : null;
        if (bmRaw && bmRaw.length) {
          const bmBase = bmRaw[0].close;
          benchmark = { symbol, series: bmRaw.map(r => ({ date: _cmpFmtDate(r.date), value: (r.close / bmBase) * 100 })) };
        }
      }

      results.push({
        ticker, market, kind, cur, currentPrice, available: true,
        holdingPeriod, samePeriod, series, drawdown, benchmark, investmentResult, markers
      });
    });

    return { success: true, period, results };
  } catch (e) {
    logError('getCompareData', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// ราคาย้อนหลัง — รวม 2 แหล่งให้อยู่ในรูปเดียวกัน [{date:Date, close, high, low}] เรียงจากเก่า→ใหม่
// ══════════════════════════════════════════════════════════
function _cmpGetOHLCSeries(ticker, market, kind) {
  if (kind === 'holding') return _getOHLCFromExternalLog(ticker); // มีอยู่แล้วใน webapp_09
  return _cmpFetchYahooDailyOHLC(ticker, market, '5y');
}

// ── ดึงราคาย้อนหลัง 5 ปีจาก Yahoo แบบมี "วันที่" กำกับทุกจุด (ต่างจาก _wlFetchYahooHistory เดิมที่ไม่มี date) ──
function _cmpFetchYahooDailyOHLC(ticker, market, rangeStr) {
  rangeStr = rangeStr || '5y';
  const symbol = (market === 'TH') ? (String(ticker).toUpperCase() + '.BK') : String(ticker).toUpperCase();
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol)
    + '?range=' + rangeStr + '&interval=1d';
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result || !result.timestamp || !result.indicators || !result.indicators.quote || !result.indicators.quote[0]) return null;

    const ts = result.timestamp;
    const q = result.indicators.quote[0];
    const rows = ts.map((t, i) => ({
      date: new Date(t * 1000),
      close: q.close ? q.close[i] : null,
      high: q.high ? q.high[i] : null,
      low: q.low ? q.low[i] : null
    })).filter(r => r.close !== null && r.close !== undefined && !isNaN(r.close));

    return rows.length ? rows : null;
  } catch (e) {
    logError('_cmpFetchYahooDailyOHLC', e);
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// Helper เล็กๆ: วันที่, หา index, rebase, drawdown, CAGR, ช่วงเวลา
// ══════════════════════════════════════════════════════════
function _cmpFmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

// หา index แรกใน series (เรียงจากเก่า→ใหม่) ที่ date >= startDate
function _cmpFindStartIndex(series, startDate) {
  for (let i = 0; i < series.length; i++) {
    if (series[i].date >= startDate) return i;
  }
  return -1;
}

function _cmpResolvePeriodStart(period) {
  const now = new Date();
  if (period === '1M') return new Date(now.getTime() - 30 * 86400000);
  if (period === '3M') return new Date(now.getTime() - 90 * 86400000);
  if (period === '6M') return new Date(now.getTime() - 182 * 86400000);
  if (period === 'YTD') return new Date(now.getFullYear(), 0, 1);
  if (period === '1Y') return new Date(now.getTime() - 365 * 86400000);
  return new Date(now.getTime() - 90 * 86400000); // default 3M
}

// rebase เป็นฐาน 100 ที่จุดเริ่มต้น + resample เป็นรายสัปดาห์ถ้าจุดข้อมูลเยอะเกินไป (> ~1 ปีเทรดดิ้ง)
// กันกราฟหนักเกินจำเป็นตอนช่วงเวลายาว (ตามคำแนะนำ: 1Y ใช้ daily, ช่วงยาวกว่านั้นใช้ weekly)
function _cmpBuildSeries(series, startIdx) {
  if (startIdx < 0 || startIdx >= series.length) return [];
  const trimmed = series.slice(startIdx);
  const base = trimmed[0].close;
  if (!base) return [];

  let points = trimmed.map(r => ({ date: _cmpFmtDate(r.date), value: (r.close / base) * 100 }));

  if (points.length > 260) {
    const weekly = [];
    for (let i = 0; i < points.length; i += 5) weekly.push(points[i]);
    if (weekly[weekly.length - 1] !== points[points.length - 1]) weekly.push(points[points.length - 1]);
    points = weekly;
  }
  return points;
}

// Drawdown % จากจุดสูงสุดที่เคยทำได้ (คำนวณจาก series ที่ rebase แล้ว)
function _cmpBuildDrawdown(points) {
  let peak = -Infinity;
  return points.map(p => {
    peak = Math.max(peak, p.value);
    return { date: p.date, value: peak > 0 ? ((p.value / peak) - 1) * 100 : 0 };
  });
}

function _cmpCAGR(returnPct, days) {
  if (!days || days < 1) return null;
  return (Math.pow(1 + returnPct / 100, 365 / days) - 1) * 100;
}

// ══════════════════════════════════════════════════════════
// Phase 2 — จุด Buy / Sell / Trailing Stop บนกราฟ
// วางจุดบน "เส้นกราฟที่ rebase แล้ว" ไม่ใช่ราคาจริง เพื่อให้จุดอยู่บนเส้นเสมอ (เหมือนแพทเทิร์น MA บนกราฟแท่งเทียน)
// ══════════════════════════════════════════════════════════

// หาจุดใน series (rebase แล้ว, เรียงเก่า→ใหม่) ที่ใกล้ที่สุดกับวันที่ที่ต้องการ (จุดแรกที่ >= วันนั้น, ถ้าเกินสุดท้ายให้ใช้จุดสุดท้าย)
// คืน null ถ้าวันที่นั้นอยู่ก่อนจุดเริ่มต้นกราฟ (เช่น ธุรกรรมเกิดก่อนช่วงเวลาที่กำลังดูอยู่ในโหมด Same Period)
function _cmpSnapToSeries(points, dateStr) {
  if (!points || !points.length || !dateStr) return null;
  if (dateStr < points[0].date) return null;
  for (let i = 0; i < points.length; i++) {
    if (points[i].date >= dateStr) return points[i];
  }
  return points[points.length - 1];
}

// ดึง TrailTierLog เป็น list พร้อมใช้ (ที่ก่อนหน้านี้มีแต่ข้อมูลดิบ ยังไม่มีตัวดึงแบบ list — เพิ่มในเฟส 2 นี้)
function _cmpGetTrailTierLogList(ticker, market) {
  try {
    const sheet = getSheet(TRAIL_LOG_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    const rows = [];
    data.forEach(row => {
      const t = row[1], m = row[2], execDate = row[4], stopPrice = row[6], triggerPrice = row[7], sharesSold = row[8], tierKey = row[3];
      if (String(t).toUpperCase() !== ticker || m !== market) return;
      if (!(execDate instanceof Date) || isNaN(execDate.getTime())) return;
      rows.push({
        dateStr: _cmpFmtDate(execDate), tierKey,
        stopPrice: Number(stopPrice) || 0, triggerPrice: Number(triggerPrice) || 0, sharesSold: Number(sharesSold) || 0
      });
    });
    rows.sort((a, b) => a.dateStr.localeCompare(b.dateStr));
    return rows;
  } catch (e) {
    logError('_cmpGetTrailTierLogList', e);
    return [];
  }
}

// รวม Buy / Sell (จาก getTickerTransactionDetail ที่ดึงมาแล้ว) + Trailing Stop trigger (จาก TrailTierLog)
// แล้ว snap แต่ละจุดลงบนเส้นกราฟที่ rebase แล้ว — คืน null ถ้าไม่มีจุดใดเลยในช่วงที่กำลังดู
function _cmpBuildMarkers(ticker, market, detail, series) {
  if (!series || !series.length) return null;
  const startDate = series[0].date, endDate = series[series.length - 1].date;

  const buys = [], sells = [];
  if (detail && detail.transactions) {
    detail.transactions.forEach(t => {
      if (t.date < startDate || t.date > endDate) return;
      const pt = _cmpSnapToSeries(series, t.date);
      if (!pt) return;
      const rec = { date: pt.date, value: pt.value, price: t.price, shares: t.shares };
      if (t.type === 'ซื้อ') buys.push(rec);
      else if (t.type === 'ขาย') sells.push(rec);
    });
  }

  const trailStops = _cmpGetTrailTierLogList(ticker, market)
    .filter(r => r.dateStr >= startDate && r.dateStr <= endDate)
    .map(r => {
      const pt = _cmpSnapToSeries(series, r.dateStr);
      if (!pt) return null;
      return { date: pt.date, value: pt.value, stopPrice: r.stopPrice, sharesSold: r.sharesSold, tierKey: r.tierKey };
    })
    .filter(Boolean);

  if (!buys.length && !sells.length && !trailStops.length) return null;
  return { buys, sells, trailStops };
}

// ══════════════════════════════════════════════════════════
// Benchmark (SPY / SET) — ดึงราคาปิดรายวันช่วงเดียวกับหุ้นที่เทียบ แล้ว rebase 100
// cache ในรอบการเรียกเดียวกัน กันดึงซ้ำถ้ามีหลายหุ้นตลาดเดียวกัน/ช่วงเวลาเดียวกัน
// หมายเหตุ: ใช้ GFINANCE_SCRATCH_SHEET ตัวเดียวกับที่ webapp_15_portfolio_history.gs ประกาศไว้แล้ว (ไม่ต้อง const ซ้ำ)
// ══════════════════════════════════════════════════════════
function _cmpGetBenchmarkSeries(market, startDate, endDate, cache) {
  const symbol = CMP_BENCHMARK_SYMBOL[market];
  if (!symbol) return null;
  const cacheKey = symbol + '|' + _cmpFmtDate(startDate);
  if (cache[cacheKey] !== undefined) return cache[cacheKey];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let scratch = ss.getSheetByName(GFINANCE_SCRATCH_SHEET);
    if (!scratch) { scratch = ss.insertSheet(GFINANCE_SCRATCH_SHEET); scratch.hideSheet(); }
    scratch.getRange('A1:B400').clearContent();

    const fmt = d => `DATE(${d.getFullYear()},${d.getMonth() + 1},${d.getDate()})`;
    scratch.getRange('A1').setFormula(
      `=GOOGLEFINANCE("${symbol}","close",${fmt(startDate)},${fmt(endDate)},"DAILY")`
    );
    SpreadsheetApp.flush();
    Utilities.sleep(1200);

    const values = scratch.getRange('A1:B400').getValues();
    scratch.getRange('A1:B400').clearContent();

    const rows = values.slice(1) // แถวแรกเป็น header (Date, Close)
      .filter(r => r[0] instanceof Date && isFinite(parseFloat(r[1])))
      .map(r => ({ date: r[0], close: parseFloat(r[1]) }));

    cache[cacheKey] = rows.length ? rows : null;
    return cache[cacheKey];
  } catch (e) {
    logError('_cmpGetBenchmarkSeries', e);
    cache[cacheKey] = null;
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// Investment Result breakdown (เฉพาะหุ้นที่ถือ) — Price Return + Dividend + FX + Fee = Total Return
// ใช้ข้อมูลจาก getTickerTransactionDetail() ที่มีอยู่แล้ว (webapp_18) ไม่ต้องดึงซ้ำ
// ══════════════════════════════════════════════════════════
function _cmpGetInvestmentResult(detail, market, currentPrice) {
  const s = detail.summary;

  // ── FX Gain/Loss (เฉพาะหุ้น US) — เทียบอัตราแลกเปลี่ยนตอนซื้อ (ถัวเฉลี่ยถ่วงน้ำหนักตามมูลค่า) กับอัตราปัจจุบัน ──
  let fxGainLossTHB = null;
  if (market === 'US') {
    const buys = detail.transactions.filter(t => t.type === 'ซื้อ' && t.exchangeRate);
    const totalAmt = buys.reduce((sum, t) => sum + t.amount, 0);
    if (totalAmt > 0) {
      const avgBuyFx = buys.reduce((sum, t) => sum + t.amount * t.exchangeRate, 0) / totalAmt;
      fxGainLossTHB = s.sharesHeld * currentPrice * (detail.fxRate - avgBuyFx);
    } else {
      fxGainLossTHB = 0;
    }
  }

  const priceReturnNative = s.unrealizedProfit + s.realizedProfit;
  const totalReturnNative = priceReturnNative + s.totalDividendNative - s.totalFee;
  const priceReturnTHB = market === 'US' ? priceReturnNative * detail.fxRate : priceReturnNative;
  const totalReturnTHB = priceReturnTHB + s.totalDividendTHB - s.totalFeeTHB + (fxGainLossTHB || 0);

  return {
    cur: detail.cur,
    capitalNative: s.buyTotal, capitalTHB: s.buyTotalTHB,
    priceReturnNative, priceReturnTHB, priceReturnPct: s.unrealizedPct,
    dividendNative: s.totalDividendNative, dividendTHB: s.totalDividendTHB,
    fxGainLossTHB, // null สำหรับหุ้นไทย (ไม่มีความเสี่ยงค่าเงิน)
    feeNative: -s.totalFee, feeTHB: -s.totalFeeTHB,
    totalReturnNative, totalReturnTHB
  };
}



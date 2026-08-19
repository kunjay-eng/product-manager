/**  webapp_22_OverAllSignal_analysis.gs
 * รวมสัญญาณจาก 6 ตัวชี้วัดที่มีอยู่แล้วในระบบ เป็นคะแนนเดียว (0-100)
 * นับ % ของสัญญาณที่เป็นขาขึ้น จากสัญญาณที่มีข้อมูลพร้อมใช้งานจริงเท่านั้น (ไม่นับตัวที่ข้อมูลไม่พอ)
 */
function getCompositeSignalScore(ticker, market) {
  try {
    const yahooSymbol = market === 'TH' ? ticker + '.BK' : ticker;
    const stockData = _fetchStockData(yahooSymbol);
    if (!stockData) return { success: false, error: 'ดึงข้อมูลราคาไม่สำเร็จ' };

    const signals = [];

    // 1) MA Trend
    signals.push({
      label: 'MA Trend (ราคา > MA20 > MA50)',
      ok: stockData.price > stockData.ma20 && stockData.ma20 > stockData.ma50,
      available: true
    });

    // 2) RSI
    signals.push({
      label: 'RSI ' + stockData.rsi.toFixed(1) + ' อยู่ในโซนแข็งแรง (50-70)',
      ok: stockData.rsi >= 50 && stockData.rsi <= 70,
      available: true
    });

    // 3) Volume
    signals.push({
      label: 'Volume สูงกว่าค่าเฉลี่ย 20 วัน',
      ok: stockData.volRatio > 1,
      available: true
    });

    // 4) MACD (อาจไม่มีถ้าข้อมูลไม่พอ)
    try {
      const ext = getExtendedTechnicals(ticker);
      if (ext.available && ext.macdLine !== undefined) {
        signals.push({ label: 'MACD Bullish (MACD Line > Signal Line)', ok: ext.macdBullish, available: true });
      }
    } catch (e) {}

    // 5) ADX
    const adx = getADXData(ticker);
    if (!adx.error) {
      signals.push({
        label: 'ADX ' + adx.adx + ' — ' + (adx.plusDI > adx.minusDI ? '+DI นำ (ขาขึ้น)' : '-DI นำ (ขาลง)'),
        ok: adx.plusDI > adx.minusDI,
        available: true
      });
    }

    // 6) Bollinger Bands
    const bb = getBollingerBands(ticker, 20);
    if (!bb.error) {
      signals.push({
        label: '%B ' + (bb.percentB * 100).toFixed(0) + '% (ครึ่งบน = โน้มเอียงขาขึ้น)',
        ok: bb.percentB > 0.5,
        available: true
      });
    }

    const total = signals.length;
    const bullishCount = signals.filter(s => s.ok).length;
    const score = total > 0 ? Math.round((bullishCount / total) * 100) : 50;

    const verdict = score >= 70 ? '🟢 สัญญาณส่วนใหญ่เป็นขาขึ้น'
      : score >= 45 ? '🟡 สัญญาณผสม ไม่ชัดเจน'
      : '🔴 สัญญาณส่วนใหญ่เป็นขาลง';

    return { success: true, ticker, score, bullishCount, total, verdict, signals };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * สแกน Holdings + Watchlist ทั้งหมด คำนวณ Composite Score ให้ทุกตัว แล้วเรียงจากสูง-ต่ำ
 */
function getScreenerData() {
  try {
    const tickerSet = {}; // key: TICKER|MARKET

    const holdings = getHoldingsData();
    (holdings.us || []).forEach(h => {
      const key = h.ticker + '|US';
      if (!tickerSet[key]) tickerSet[key] = { ticker: h.ticker, market: 'US', sources: [] };
      tickerSet[key].sources.push('Holdings');
    });
    (holdings.th || []).forEach(h => {
      const key = h.ticker + '|TH';
      if (!tickerSet[key]) tickerSet[key] = { ticker: h.ticker, market: 'TH', sources: [] };
      tickerSet[key].sources.push('Holdings');
    });

    const wl = getWatchlistData();
    (wl.items || []).forEach(w => {
      const key = w.ticker + '|' + w.market;
      if (!tickerSet[key]) tickerSet[key] = { ticker: w.ticker, market: w.market, sources: [] };
      if (tickerSet[key].sources.indexOf('Watchlist') === -1) tickerSet[key].sources.push('Watchlist');
    });

    const list = Object.values(tickerSet);
    if (list.length === 0) return { success: false, error: 'ไม่มีหุ้นใน Holdings หรือ Watchlist เลย' };

    const results = [];
    list.forEach(t => {
      try {
        const r = getCompositeSignalScore(t.ticker, t.market);
        if (r.success) {
          results.push({ ticker: t.ticker, market: t.market, sources: t.sources, score: r.score, verdict: r.verdict, bullishCount: r.bullishCount, total: r.total });
        } else {
          results.push({ ticker: t.ticker, market: t.market, sources: t.sources, error: r.error });
        }
      } catch (e) {
        results.push({ ticker: t.ticker, market: t.market, sources: t.sources, error: e.message });
      }
    });

    results.sort((a, b) => (b.score ?? -1) - (a.score ?? -1));

    return { success: true, count: results.length, results };
  } catch (e) {
    return { success: false, error: e.message };
  }
}



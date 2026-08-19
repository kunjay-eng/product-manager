// ══════════════════════════════════════════════════════════
// webapp_17_relative_strength.gs — RS Line (Relative Strength vs Benchmark)
// ------------------------------------------------------------
// RS = ราคาหุ้น / ราคา Benchmark ณ วันเดียวกัน แล้ว index ให้เริ่มที่ 100
// เส้น RS ขาขึ้น = หุ้นวิ่งแรงกว่าตลาด (แม้ราคาหุ้นจะขึ้น/ลงก็ตาม)
// เส้น RS ขาลง  = หุ้นวิ่งอ่อนกว่าตลาด
//
// ⚠️ ใช้ BENCHMARK_SYMBOLS / GFINANCE_SCRATCH_SHEET ร่วมกับ
//    webapp_15_portfolio_history.gs (ประกาศไว้แล้ว ไม่ต้องประกาศซ้ำ)
// ⚠️ หนักกว่า indicator อื่นเพราะดึงราคาย้อนหลังของ index ทั้งช่วง —
//    ตั้งใจไม่ผูกกับ auto-refresh 30 วิ ให้โหลดครั้งเดียว/กดรีเฟรชเอง
// ══════════════════════════════════════════════════════════

const RS_LOOKBACK_DAYS = 130; // ~6 เดือนเทรดดิ้ง
const RS_MIN_DAYS = 40;       // ขั้นต่ำถึงจะเชื่อถือได้
const RS_TREND_LOOKBACK = 20; // เทียบ RS ปัจจุบัน vs 20 วันก่อน

function getRelativeStrengthData(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const stockHist = _getOHLCFromExternalLog(ticker); // reuse จาก webapp_09
    if (!stockHist) {
      return { available: false, reason: 'ไม่พบข้อมูลราคาย้อนหลังของ ' + ticker + ' ใน Daily_Close_Log' };
    }
    if (stockHist.length < RS_MIN_DAYS) {
      return { available: false, reason: `มีข้อมูลราคาแค่ ${stockHist.length} วัน (ต้องการอย่างน้อย ${RS_MIN_DAYS} วัน) — รอสะสมข้อมูลเพิ่ม` };
    }

    const stockSlice = stockHist.slice(-RS_LOOKBACK_DAYS);
    const startDate = new Date(stockSlice[0].date);
    startDate.setDate(startDate.getDate() - 5); // กันชนวันเปิดตลาดไม่ตรงกัน
    const endDate = new Date();

    const benchSymbol = (market === 'TH') ? BENCHMARK_SYMBOLS.SET : BENCHMARK_SYMBOLS.SP500;
    const benchName = (market === 'TH') ? 'SET Index' : 'S&P 500';

    const benchSeries = _getGoogleFinanceHistoricalSeries(benchSymbol, startDate, endDate);
    if (!benchSeries || benchSeries.length < RS_MIN_DAYS) {
      return { available: false, reason: `ดึงราคาย้อนหลังของ ${benchName} ไม่สำเร็จ หรือได้ข้อมูลไม่พอ (GOOGLEFINANCE อาจช้า/ติด rate limit ลองใหม่อีกครั้ง)` };
    }

    // ── map benchmark ตามวันที่ (yyyy-MM-dd) เพื่อจับคู่กับราคาหุ้น ──
    const benchMap = {};
    benchSeries.forEach(b => { benchMap[Utilities.formatDate(b.date, 'Asia/Bangkok', 'yyyy-MM-dd')] = b.close; });

    // ── จับคู่ตามวันที่จริง + carry-forward ถ้าวันหยุดตลาดไม่ตรงกัน ──
    const aligned = [];
    let lastBench = null;
    stockSlice.forEach(s => {
      const key = Utilities.formatDate(new Date(s.date), 'Asia/Bangkok', 'yyyy-MM-dd');
      if (benchMap[key] !== undefined) lastBench = benchMap[key];
      if (lastBench !== null && s.close > 0) {
        aligned.push({ date: s.date, stockClose: s.close, benchClose: lastBench });
      }
    });

    if (aligned.length < RS_MIN_DAYS) {
      return { available: false, reason: 'จับคู่วันที่ระหว่างหุ้นกับ Benchmark ได้ไม่พอ (มี ' + aligned.length + ' วัน)' };
    }

    // ── RS Ratio → Index เริ่มที่ 100 ──
    const rsRatioStart = aligned[0].stockClose / aligned[0].benchClose;
    const rsIndexed = aligned.map(a => ((a.stockClose / a.benchClose) / rsRatioStart) * 100);

    const n = rsIndexed.length;
    const lookback = Math.min(RS_TREND_LOOKBACK, n - 1);
    const rsNow = rsIndexed[n - 1];
    const rsBefore = rsIndexed[n - 1 - lookback];
    const changePct = ((rsNow - rsBefore) / rsBefore) * 100;

    // ── slope เส้น RS ช่วง lookback หลังสุด (least squares) — กันสัญญาณหลอกจาก 2 จุดเดียว ──
    const tail = rsIndexed.slice(-Math.min(RS_TREND_LOOKBACK, n));
    const slope = _rsLinearSlope(tail);

    let trend, trendClass, trendLabel;
    if (changePct > 2 && slope > 0) {
      trend = 'up'; trendClass = 'safe'; trendLabel = '📈 RS ขาขึ้น — หุ้นแข็งแกร่งกว่าตลาด';
    } else if (changePct < -2 && slope < 0) {
      trend = 'down'; trendClass = 'stop'; trendLabel = '📉 RS ขาลง — หุ้นอ่อนแอกว่าตลาด';
    } else {
      trend = 'flat'; trendClass = 'warn'; trendLabel = '➖ RS แกว่งตัว — ยังไม่ชัดเจนว่าแรงกว่า/อ่อนกว่าตลาด';
    }

    return {
      available: true,
      ticker, market,
      benchmarkName: benchName,
      rsNow: _taxRound(rsNow, 2),
      changePct: _taxRound(changePct, 2),
      lookbackDays: lookback,
      trend, trendClass, trendLabel,
      sparklineValues: rsIndexed.slice(-60), // ให้ frontend ใช้ buildSparkline() เดิม
      daysAnalyzed: n,
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('getRelativeStrengthData', e);
    return { available: false, reason: e.message };
  }
}

// ── ดึงราคาปิดย้อนหลังทั้งช่วงจาก GOOGLEFINANCE (1 ครั้ง คืนทั้ง series)
//    ใช้ scratch sheet เดียวกับ webapp_15 — หนักกว่าดึงจุดเดียว จึง retry นานกว่าปกติ ──
function _getGoogleFinanceHistoricalSeries(symbol, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let scratch = ss.getSheetByName(GFINANCE_SCRATCH_SHEET);
  if (!scratch) { scratch = ss.insertSheet(GFINANCE_SCRATCH_SHEET); scratch.hideSheet(); }
  scratch.getRange('A1:B260').clearContent();

  const fmt = d => `DATE(${d.getFullYear()},${d.getMonth() + 1},${d.getDate()})`;
  scratch.getRange('A1').setFormula(
    `=GOOGLEFINANCE("${symbol}","close",${fmt(startDate)},${fmt(endDate)},"DAILY")`
  );
  SpreadsheetApp.flush();

  let values = null;
  for (let i = 0; i < 10; i++) {
    Utilities.sleep(700);
    values = scratch.getRange('A1:B260').getValues();
    const dataRowCount = values.filter(r => r[0] instanceof Date).length;
    if (dataRowCount >= RS_MIN_DAYS) break;
  }
  scratch.getRange('A1:B260').clearContent();

  if (!values) return null;
  const series = values
    .filter(r => r[0] instanceof Date && isFinite(parseFloat(r[1])) && parseFloat(r[1]) > 0)
    .map(r => ({ date: r[0], close: parseFloat(r[1]) }))
    .sort((a, b) => a.date - b.date);

  return series.length ? series : null;
}

// ── Least-squares slope ของ series (แกน x = index 0..n-1) ──
function _rsLinearSlope(values) {
  const n = values.length;
  if (n < 2) return 0;
  const xMean = (n - 1) / 2;
  const yMean = values.reduce((s, v) => s + v, 0) / n;
  let num = 0, den = 0;
  for (let i = 0; i < n; i++) {
    num += (i - xMean) * (values[i] - yMean);
    den += (i - xMean) * (i - xMean);
  }
  return den !== 0 ? num / den : 0;
}

// ============================================================
// webapp_03_analyze.gs — หน้าวิเคราะห์ (รวม Portfolio-mode + Fast-mode)
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
//
// ส่วนที่ 1: getPortfolioData() — รายชื่อหุ้นใน ATR Portfolio (ใช้เป็น
//            ลิสต์เลือกหุ้นในหน้าวิเคราะห์ ทั้งสองโหมด)
// ส่วนที่ 2: getStockAnalysis() — Portfolio-mode deep dive (MA/RSI/Volume
//            + Decision Engine เดียวกับ /analyze ของ Telegram Bot เดิม)
// ส่วนที่ 3: getFastSignal()/getFastSignalList() — Fast-mode (EMA5/EMA20
//            cross + RSI + Volume แยก Trend ออกจาก Position Management
//            ใช้ getEffectiveRiskParams() จาก webapp_05_settings.gs)
// ============================================================

// ────────────────────────────────────────────────
// ส่วนที่ 1: รายชื่อหุ้นใน ATR Portfolio
// ────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════
// ATR Portfolio Dashboard — getPortfolioData()
// ✅ ใช้ _getATRRows() ตัวเดียวกับที่ atr_report.gs ใช้อยู่แล้ว
//    (ไม่เขียน readATRSheet ซ้ำ — เลี่ยง column mapping ขัดกัน)
// ══════════════════════════════════════════════════════════
function getPortfolioData() {
  try {
    return {
      us: _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US),
      th: _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH)
    };
  } catch (e) {
    logError('getPortfolioData', e);
    return { us: [], th: [], error: e.message };
  }
}


// ────────────────────────────────────────────────
// ส่วนที่ 2: Portfolio-mode deep dive analysis
// ────────────────────────────────────────────────
function getStockAnalysis(ticker) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'กรุณาระบุชื่อหุ้น' };

    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    let atr = usRows.find(r => r.symbol.toUpperCase() === ticker);
    let isTH = false;
    if (!atr) {
      atr = thRows.find(r => r.symbol.toUpperCase() === ticker);
      isTH = !!atr;
    }
    if (!atr) {
      return {
        success: false,
        error: 'ไม่พบ ' + ticker + ' ใน ATR Portfolio',
        available: { us: usRows.map(r => r.symbol), th: thRows.map(r => r.symbol) }
      };
    }

    const yahooSymbol = isTH ? ticker + '.BK' : ticker;
    const stockData = _fetchStockData(yahooSymbol);
    if (!stockData) return { success: false, error: 'ดึงข้อมูล Yahoo ไม่ได้: ' + ticker };

    const cur   = isTH ? '฿' : '$';
    const price = atr.priceNow > 0 ? atr.priceNow : stockData.price;
    const ma20  = stockData.ma20, ma50 = stockData.ma50, ma200 = stockData.ma200;
    const rsi   = stockData.rsi;
    const volNow = stockData.volNow, volAvg = stockData.volAvg20, volRatio = stockData.volRatio;

    const plPct = atr.buyPrice > 0 ? ((price - atr.buyPrice) / atr.buyPrice) * 100 : 0;
    const belowStop = price < atr.trailingStop;
    const distToStop = atr.trailingStop > 0 ? ((price - atr.trailingStop) / price) * 100 : 0;

    // ── 2. P/L Assessment ──
    let plSignal = '', plClass = '';
    if (plPct > 20)       { plSignal = 'กำไรดี — เริ่มปกป้องกำไร';    plClass = 'safe'; }
    else if (plPct > 0)   { plSignal = 'กำไรน้อย — ดูแนวโน้มต่อ';      plClass = 'warn'; }
    else if (plPct > -10) { plSignal = 'ติดลบเล็กน้อย — เฝ้าระวัง';    plClass = 'warn'; }
    else if (plPct > -15) { plSignal = 'ติดลบ >10% — ระวัง';           plClass = 'stop'; }
    else                  { plSignal = 'ติดลบ >15% — พิจารณาตัด';      plClass = 'stop'; }

    // ── 3. Trend ──
    let trendSignal = '', trendClass = '';
    if (price > ma20 && ma20 > ma50 && ma50 > ma200) {
      trendSignal = 'ขาขึ้นแข็งแรง (Price>MA20>MA50>MA200)'; trendClass = 'safe';
    } else if (price > ma50 && ma50 > ma200) {
      trendSignal = 'ขาขึ้น แต่ราคาต่ำกว่า MA20'; trendClass = 'warn';
    } else if (price < ma20 && price < ma50) {
      trendSignal = 'ขาลง (ราคาต่ำกว่า MA20 และ MA50)'; trendClass = 'stop';
    } else {
      trendSignal = 'แนวโน้มผสม'; trendClass = 'warn';
    }

    // ── 4. RSI ──
    let rsiSignal = '', rsiClass = '';
    if (rsi > 70)        { rsiSignal = 'ร้อนแรงเกินไป'; rsiClass = 'stop'; }
    else if (rsi > 50)   { rsiSignal = 'แข็งแรง';         rsiClass = 'safe'; }
    else if (rsi > 30)   { rsiSignal = 'อ่อนตัว';         rsiClass = 'warn'; }
    else                 { rsiSignal = 'Oversold';        rsiClass = 'safe'; }

    // ── 5. Volume ──
    let volSignal = '', volClass = '';
    if (plPct < 0 && volRatio > 1.5) {
      volSignal = 'ราคาลง + Volume สูง ×' + fmt(volRatio) + ' — สัญญาณลบ'; volClass = 'stop';
    } else if (plPct > 0 && volRatio > 1.5) {
      volSignal = 'ราคาขึ้น + Volume สูง ×' + fmt(volRatio) + ' — สัญญาณบวก'; volClass = 'safe';
    } else {
      volSignal = 'Volume ปกติ ×' + fmt(volRatio); volClass = 'warn';
    }

    // ── สรุป Decision (ตรรกะเดียวกับ cmdAnalyze เดิม) ──
    const reasons = [], warnings = [];
    let decision = '', decClass = '', plan = '';

    if (belowStop) {
      decision = 'ขายทันที'; decClass = 'stop';
      reasons.push('ราคาต่ำกว่า Trailing Stop ' + cur + fmt(atr.trailingStop));
      reasons.push('ระบบบอกว่าแนวโน้มเสียแล้ว');
      plan = 'ขายออกทั้งหมดทันที\nStop: ' + cur + fmt(atr.trailingStop) +
             '\nราคาปัจจุบัน: ' + cur + fmt(price);

    } else if (plPct < -15) {
      decision = 'พิจารณาตัดขาดทุน'; decClass = 'stop';
      reasons.push('ขาดทุน ' + fmt(plPct) + '% เกิน 15%');
      if (price < ma50) reasons.push('ราคาต่ำกว่า MA50');
      plan = 'พิจารณาตัดขาดทุน\nถ้าราคาไม่ฟื้นกลับเหนือ ' + cur + fmt(ma50) +
             ' (MA50)\nให้ขายออกเพื่อจำกัดความเสียหาย';

    } else if (price > ma20 && rsi < 70 && !belowStop && plPct >= 0) {
      decision = 'ถือ'; decClass = 'safe';
      if (price > ma20) reasons.push('ราคายืนเหนือ MA20');
      if (ma20 > ma50)  reasons.push('MA20 > MA50');
      if (ma50 > ma200) reasons.push('MA50 > MA200');
      if (volRatio > 1.5 && plPct > 0) reasons.push('Volume สนับสนุนการขึ้น');
      if (rsi > 50 && rsi < 70) reasons.push('RSI แข็งแรง');
      if (rsi >= 65) warnings.push('RSI เริ่มสูง ติดตามใกล้ชิด');
      plan = 'ถือรันกำไรต่อ\nหากราคาปิดต่ำกว่า ' + cur + fmt(atr.trailingStop) + ' ให้ขายทันที';

    } else if (plPct < 0 && price > ma50) {
      decision = 'ถือ (เฝ้าระวัง)'; decClass = 'warn';
      reasons.push('ราคายังอยู่เหนือ MA50');
      reasons.push('ราคาต่ำกว่า MA20 — แรงขายระยะสั้น');
      if (rsi < 40) reasons.push('RSI ต่ำ อาจดีดกลับ');
      plan = 'เฝ้าระวัง\nรอราคากลับขึ้นเหนือ MA20 (' + cur + fmt(ma20) + ')' +
             '\nStop: ' + cur + fmt(atr.trailingStop);

    } else {
      decision = 'เฝ้าระวัง'; decClass = 'warn';
      reasons.push('สัญญาณไม่ชัดเจน');
      plan = 'ติดตามใกล้ชิด\nStop: ' + cur + fmt(atr.trailingStop);
    }

    const updatedAt = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

    return {
      success: true,
      isTH: isTH,
      symbol: atr.symbol,
      tradeStyle: atr.tradeStyle,
      cur: cur,
      buyPrice: atr.buyPrice,
      highestClose: atr.highestClose,
      price: price,
      atrVal: atr.atr,
      multiplier: atr.multiplier,
      trailingStop: atr.trailingStop,
      distToStop: distToStop,
      belowStop: belowStop,
      plPct: plPct, plSignal: plSignal, plClass: plClass,
      ma20: ma20, ma50: ma50, ma200: ma200,
      trendSignal: trendSignal, trendClass: trendClass,
      rsi: rsi, rsiSignal: rsiSignal, rsiClass: rsiClass,
      volNow: volNow, volAvg: volAvg, volRatio: volRatio,
      volSignal: volSignal, volClass: volClass,
      decision: decision, decClass: decClass,
      reasons: reasons, warnings: warnings, plan: plan,
      updatedAt: updatedAt
    };
  } catch (e) {
    logError('getStockAnalysis', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────
// ส่วนที่ 3: Fast Trade Signal (EMA/RSI/Volume + Position)
// ────────────────────────────────────────────────
const FAST_HISTORY_DAYS = 60; // ดึงย้อนหลัง ~60 วันทำการ พอสำหรับ EMA20 ที่นิ่งแล้ว
const ATR_SHARES_COL = 11; // col K — "จำนวนหุ้น" ใน 📊 ATR_Portfolio US/TH


// ── อ่านจำนวนหุ้นจากคอลัมน์ K ตรงๆ (แยกจาก _getATRRows เพราะไม่รู้ว่าตัวนั้น
//    ดึงคอลัมน์นี้ออกมาด้วยหรือเปล่า — อ่านเองปลอดภัยกว่าเดาแก้ฟังก์ชันที่มองไม่เห็น) ──
function _getATRShares(sheetName, startRow, ticker) {
  try {
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return null;
    const numRows = lastRow - startRow + 1;
    const symbols = sheet.getRange(startRow, 1, numRows, 1).getValues();
    const rowIdx = symbols.findIndex(r => String(r[0]).trim().toUpperCase() === ticker);
    if (rowIdx === -1) return null;
    const val = sheet.getRange(startRow + rowIdx, ATR_SHARES_COL).getValue();
    const num = parseFloat(val);
    return isNaN(num) ? null : num;
  } catch (e) {
    logError('_getATRShares', e);
    return null;
  }
}


// ══════════════════════════════════════════════════════════
// EMA แบบ series (คืนค่าทุกจุด ไม่ใช่แค่ตัวล่าสุด) เพื่อเช็คจุดตัดกัน
// (จุดก่อน period-1 ตัวแรกเป็น null เพราะข้อมูลไม่พอ)
// ══════════════════════════════════════════════════════════
function _calcEMASeries(closes, period) {
  const k = 2 / (period + 1);
  const out = [];
  let ema = null;
  for (let i = 0; i < closes.length; i++) {
    if (i < period - 1) { out.push(null); continue; }
    if (i === period - 1) {
      let sum = 0;
      for (let j = 0; j < period; j++) sum += closes[j];
      ema = sum / period;
    } else {
      ema = closes[i] * k + ema * (1 - k);
    }
    out.push(ema);
  }
  return out;
}

// RSI(14) แบบ simple average (เพียงพอสำหรับสัญญาณระดับนี้)
function _calcRSI(closes, period) {
  period = period || 14;
  if (closes.length < period + 1) return 50; // ข้อมูลไม่พอ → คืนค่ากลาง กันหน้าจอพัง
  let gains = 0, losses = 0;
  const start = closes.length - period;
  for (let i = start; i < closes.length; i++) {
    const diff = closes[i] - closes[i - 1];
    if (diff >= 0) gains += diff; else losses -= diff;
  }
  const avgGain = gains / period;
  const avgLoss = losses / period;
  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - (100 / (1 + rs));
}

// ══════════════════════════════════════════════════════════
// วิเคราะห์หุ้น 1 ตัวแบบ Fast Signal — เรียกจากหน้าเว็บผ่าน
// google.script.run.getFastSignal(ticker)
// ══════════════════════════════════════════════════════════
function getFastSignal(ticker) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'กรุณาระบุชื่อหุ้น' };

    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    let atr = usRows.find(r => r.symbol.toUpperCase() === ticker);
    let isTH = false;
    if (!atr) {
      atr = thRows.find(r => r.symbol.toUpperCase() === ticker);
      isTH = !!atr;
    }
    if (!atr) {
      return { success: false, error: 'ไม่พบ ' + ticker + ' ใน ATR Portfolio' };
    }

    const yahooSymbol = isTH ? ticker + '.BK' : ticker;
    const history = fetchYahooHistory(yahooSymbol, FAST_HISTORY_DAYS);
    if (!history || history.length < 21) {
      return { success: false, error: 'ดึงข้อมูลราคาย้อนหลังไม่พอสำหรับคำนวณ EMA20: ' + ticker };
    }

    //const closes = history.map(h => h.close);
    //const vols   = history.map(h => h.volume);
    //const n = closes.length;
    //const price = atr.priceNow > 0 ? atr.priceNow : closes[n - 1];
    //const cur = isTH ? '฿' : '$';

    const closes = history.map(h => h.close);
    const vols   = history.map(h => h.volume);
    const n = closes.length;

      // ── ดึงราคาสดจริงผ่าน lookupTickerPrice() แทนอ่าน atr.priceNow จากชีต ATR Portfolio ──
    // (เซลล์ในชีตอัปเดตช้ากว่านี้มาก ทำให้ปุ่ม ▶️ auto-refresh ทุก 30 วิ เห็นราคาเดิมซ้ำ)
    // ⚠️ เขียนกลับเข้า Holdings sheet เท่านั้น (ไม่ใช่ ATR sheet โดยตรง) เพราะคอลัมน์ PRICE_NOW
    //    ในชีต ATR เป็น "สูตร" ที่ดึงค่าจาก Holdings มาอีกที — ถ้าเขียนทับ ATR ตรงๆ จะลบสูตรทิ้งถาวร
    //    เขียนเข้า Holdings แล้วปล่อยให้สูตรในชีต ATR ดึงค่าใหม่ตามมาเองอัตโนมัติ
    let price;
    try {
      const live = lookupTickerPrice(ticker, isTH ? 'TH' : 'US');
      if (live && live.success && live.price > 0) {
        price = live.price;
        const sheetName = isTH ? SHEETS.TH_HOLD : SHEETS.US_HOLD;
        writeLivePriceToHoldingsSheet(sheetName, ticker, live.price); // ← เขียนเข้า Holdings เท่านั้น
      } else {
        price = atr.priceNow > 0 ? atr.priceNow : closes[n - 1];
      }
    } catch (e) {
      price = atr.priceNow > 0 ? atr.priceNow : closes[n - 1]; // fallback กันพัง ถ้า lookupTickerPrice error
    }

    const cur = isTH ? '฿' : '$';


    // ── Trend: EMA5/EMA20 cross + RSI + Volume ──
    const ema5Arr  = _calcEMASeries(closes, 5);
    const ema20Arr = _calcEMASeries(closes, 20);
    const ema5Now  = ema5Arr[n - 1],  ema20Now  = ema20Arr[n - 1];
    const ema5Prev = ema5Arr[n - 2],  ema20Prev = ema20Arr[n - 2];
    const rsi = _calcRSI(closes, 14);

    const volToday = vols[n - 1];
    const volHist  = vols.slice(Math.max(0, n - 21), n - 1); // 20 วันก่อนหน้า (ไม่รวมวันนี้)
    const volAvg20 = volHist.length ? volHist.reduce((a, b) => a + b, 0) / volHist.length : volToday;
    const volRatio = volAvg20 > 0 ? volToday / volAvg20 : 1;

    let crossType = 'none';
    if (ema5Prev !== null && ema20Prev !== null) {
      if (ema5Prev <= ema20Prev && ema5Now > ema20Now) crossType = 'golden';
      else if (ema5Prev >= ema20Prev && ema5Now < ema20Now) crossType = 'death';
      else if (ema5Now > ema20Now) crossType = 'bullish';
      else crossType = 'bearish';
    }

    let trendStatus = 'neutral', trendClass = 'warn', trendLabel = 'Sideways / รอสัญญาณ';
    if (crossType === 'golden' && rsi > 50 && rsi < 75 && volRatio > 1.2) {
      trendStatus = 'bullish'; trendClass = 'safe'; trendLabel = 'Strong Buy — Golden Cross + Volume ยืนยัน';
    } else if (crossType === 'golden' || (crossType === 'bullish' && rsi >= 50)) {
      trendStatus = 'bullish'; trendClass = 'safe'; trendLabel = 'Uptrend';
    } else if (crossType === 'death' || (crossType === 'bearish' && rsi < 50)) {
      trendStatus = 'bearish'; trendClass = 'stop'; trendLabel = 'Downtrend';
    } else {
      trendStatus = 'neutral'; trendClass = 'warn'; trendLabel = 'Sideways / รอสัญญาณ';
    }

    const emaConditions = [
      { cls: (ema5Now > ema20Now) ? 'safe' : 'stop',
        text: 'EMA5 ' + (ema5Now > ema20Now ? '>' : '<') + ' EMA20' +
              (crossType === 'golden' ? ' (เพิ่งตัดขึ้น 🟢)' : crossType === 'death' ? ' (เพิ่งตัดลง 🔴)' : '') },
      { cls: rsi > 70 ? 'stop' : (rsi >= 50 ? 'safe' : (rsi >= 30 ? 'warn' : 'safe')),
        text: 'RSI ' + rsi.toFixed(2) + (rsi > 70 ? ' (ร้อนแรงเกินไป)' : rsi >= 50 ? ' (แข็งแรง)' : rsi >= 30 ? ' (อ่อนตัว)' : ' (Oversold)') },
      { cls: volRatio > 1.5 ? (price >= closes[n - 2] ? 'safe' : 'stop') : 'warn',
        text: 'Volume ×' + volRatio.toFixed(2) + (volRatio > 1.5 ? (price >= closes[n - 2] ? ' (สูง หนุนขาขึ้น)' : ' (สูง กดดันขาลง)') : ' (ปกติ)') }
    ];

    // ── Position: Cut Stop (จากต้นทุน) + Trailing Stop (จาก highest close เดิม) ──
    const eff = getEffectiveRiskParams(ticker);
    const cutStopAtrX   = eff.cutStopAtrX   !== null ? eff.cutStopAtrX   : atr.multiplier;
    const trailStopAtrX = eff.trailStopAtrX !== null ? eff.trailStopAtrX : atr.multiplier;

    const cutStopPrice      = atr.buyPrice - atr.atr * cutStopAtrX;
    const trailingStopPrice = atr.trailingStop; // ของเดิมจาก ATR_Portfolio (highestClose - ATR*multiplier)
    const stopLine = Math.max(cutStopPrice, trailingStopPrice);
    const belowStop = price <= stopLine;

    const plPct = atr.buyPrice > 0 ? ((price - atr.buyPrice) / atr.buyPrice) * 100 : 0;

    const shares = _getATRShares(isTH ? ATR_SHEETS.TH : ATR_SHEETS.US, isTH ? ATR_START_ROW.TH : ATR_START_ROW.US, ticker); // ← เพิ่มบรรทัดนี้
   

    let positionStatus = 'holding', positionClass = 'safe', positionLabel = 'Holding — ปกติ';
    if (belowStop) {
      positionStatus = 'exit'; positionClass = 'stop'; positionLabel = 'Exit Triggered — หลุด Stop';
    } else if (plPct >= eff.takeProfitPct * 100) {
      positionStatus = 'lock'; positionClass = 'warn'; positionLabel = 'พิจารณาล็อกกำไร';
    }

    // ── Decision Engine: รวม Trend + Position (Position มาก่อนเสมอถ้าหลุด Stop) ──
    const reasonsFor = [], reasonsAgainst = [];
    let decision, decisionClass, decisionNote, actionPlan = null;

    if (belowStop) {
      decision = 'EXIT'; decisionClass = 'stop';
      if (trendStatus === 'bullish') reasonsFor.push('แนวโน้มหลัก (EMA) ยังเป็นขาขึ้น');
      else if (trendStatus === 'bearish') reasonsAgainst.push('แนวโน้มหลักเริ่มกลับเป็นขาลงด้วย');
      reasonsAgainst.push('ราคาปิดต่ำกว่า Stop (' + cur + fmt(stopLine) + ')');
      decisionNote = trendStatus === 'bullish'
        ? 'ระบบ Fast Trade ถือว่าหุ้นเสีย Momentum ระยะสั้น แม้แนวโน้มหลักยังดี — ให้ยึดวินัย Stop ก่อน'
        : 'ทั้งแนวโน้มและสถานะเสียพร้อมกัน — ควรออกจากสถานะ';
      actionPlan = {
        action: 'ขายทั้งหมด',
        expected: plPct >= 0 ? 'ล็อกกำไรที่ทำได้' : 'ลดความเสี่ยง/จำกัดขาดทุน',
        waitConditions: ['ราคากลับขึ้นเหนือ EMA20 (' + cur + fmt(ema20Now) + ')', 'เกิดสัญญาณ Golden Cross ใหม่']
      };
    } else if (positionStatus === 'lock') {
      decision = 'WATCH'; decisionClass = 'warn';
      reasonsFor.push('ยังไม่หลุด Stop (' + cur + fmt(stopLine) + ')');
      if (trendStatus === 'bullish') reasonsFor.push('แนวโน้มหลักยังเป็นขาขึ้น');
      reasonsAgainst.push('กำไรเกิน ' + (eff.takeProfitPct * 100).toFixed(0) + '% ควรพิจารณาล็อกกำไรบางส่วน');
      decisionNote = 'ถือต่อได้ แต่เริ่มพิจารณาล็อกกำไรบางส่วนตามแผน';
    } else if (trendStatus === 'bullish') {
      decision = 'HOLD'; decisionClass = 'safe';
      reasonsFor.push('EMA5 > EMA20 — แนวโน้มขาขึ้น');
      reasonsFor.push('ยังไม่หลุด Stop (' + cur + fmt(stopLine) + ')');
      decisionNote = 'แนวโน้มและสถานะไปด้วยกัน — ถือต่อตามระบบ';
    } else if (trendStatus === 'bearish') {
      decision = 'WATCH'; decisionClass = 'warn';
      reasonsFor.push('ยังไม่หลุด Stop (' + cur + fmt(stopLine) + ')');
      reasonsAgainst.push('แนวโน้มเริ่มอ่อนตัว (EMA5 < EMA20)');
      decisionNote = 'สถานะยังไม่เสีย แต่แนวโน้มเริ่มไม่สนับสนุน — เฝ้าระวังใกล้ชิด อาจขยับ Stop ให้ตึงขึ้น';
    } else {
      decision = 'WATCH'; decisionClass = 'warn';
      reasonsFor.push('ยังไม่หลุด Stop (' + cur + fmt(stopLine) + ')');
      decisionNote = 'สัญญาณยังไม่ชัดเจน รอ EMA ตัดกันหรือ RSI/Volume ยืนยันทิศทาง';
    }

    // ── คะแนนดาว (0-5) ──
    let momentumStars = 3;
    if (rsi > 70) momentumStars = 4;
    else if (rsi >= 50) momentumStars = 5;
    else if (rsi >= 30) momentumStars = 3;
    else momentumStars = 2;
    if (volRatio > 1.5) momentumStars += (price >= closes[n - 2] ? 1 : -1);
    momentumStars = Math.max(1, Math.min(5, momentumStars));

    let riskStars;
    if (belowStop) riskStars = 1;
    else {
      const distPct = ((price - stopLine) / price) * 100;
      riskStars = distPct > 15 ? 5 : distPct > 10 ? 4 : distPct > 5 ? 3 : distPct > 0 ? 2 : 1;
    }

    const trendStars = trendStatus === 'bullish' ? 5 : trendStatus === 'neutral' ? 3 : 1;
    const positionVerdict = decision === 'EXIT' ? 'SELL' : decision === 'HOLD' ? 'HOLD' : 'WATCH';

    return {
      success: true,
      isTH: isTH, symbol: atr.symbol, tradeStyle: atr.tradeStyle, cur: cur,
      price: price, buyPrice: atr.buyPrice, highestClose: atr.highestClose, plPct: plPct, shares: shares, // ← เพิ่ม

      trendStatus: trendStatus, trendClass: trendClass, trendLabel: trendLabel,
      emaConditions: emaConditions, rsi: rsi, volRatio: volRatio,
      ema5: ema5Now, ema20: ema20Now,

      positionStatus: positionStatus, positionClass: positionClass, positionLabel: positionLabel,
      atrVal: atr.atr, cutStopAtrX: cutStopAtrX, trailStopAtrX: trailStopAtrX,
      cutStopPrice: cutStopPrice, trailingStopPrice: trailingStopPrice, stopLine: stopLine, belowStop: belowStop,

      decision: decision, decisionClass: decisionClass, decisionNote: decisionNote,
      reasonsFor: reasonsFor, reasonsAgainst: reasonsAgainst, actionPlan: actionPlan,

      trendStars: trendStars, momentumStars: momentumStars, riskStars: riskStars, positionVerdict: positionVerdict,

      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
    };
  } catch (e) {
    logError('getFastSignal', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// รายการหุ้นโหมด Fast ทั้งหมด พร้อมสัญญาณสรุปสั้นๆ — ใช้แสดงลิสต์เลือก
// หุ้นในหน้าเว็บ (เร็วกว่าเรียก getFastSignal() ทีละตัว เพราะไม่ต้อง
// เขียนซ้ำ logic แต่ยังคง fetch ราคาย้อนหลังทีละตัวเช่นเดิม — ถ้าหุ้นเยอะ
// มากในอนาคตค่อยพิจารณาแคช)
// ══════════════════════════════════════════════════════════
function getFastSignalList() {
  try {
    const modeMap = getStockModeMap();
    const fastTickers = Object.keys(modeMap).filter(t => modeMap[t].mode === 'Fast');

    const results = fastTickers.map(t => {
      const r = getFastSignal(t);
      if (!r.success) return { symbol: t, error: r.error };
      return {
        symbol: r.symbol, isTH: r.isTH, price: r.price, plPct: r.plPct,
        trendClass: r.trendClass, positionClass: r.positionClass,
        decision: r.decision, decisionClass: r.decisionClass
      };
    });

    return { success: true, data: results };
  } catch (e) {
    logError('getFastSignalList', e);
    return { success: false, error: e.message, data: [] };
  }
}



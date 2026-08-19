/* ============================================================
   ## webapp_21_Bollinger_analysis.gs
   📊 Bollinger Bands
   ============================================================ */

/**
 * คำนวณ Bollinger Bands ทั้งซีรีส์ จาก OHLC (ต้องมี field .close ในแต่ละวัน)
 * period ปกติ 20 วัน, stdDevMultiplier ปกติ 2 เท่า
 * คืนค่าเป็น array {date, sma, upper, lower, percentB, bandwidth} ทุกวันที่คำนวณได้
 */
function _calculateBollingerBands(ohlc, period, stdDevMultiplier) {
  period = period || 20;
  stdDevMultiplier = stdDevMultiplier || 2;
  if (!ohlc || ohlc.length < period) return [];

  const result = [];
  for (let i = period - 1; i < ohlc.length; i++) {
    const slice = ohlc.slice(i - period + 1, i + 1).map(o => o.close);
    const sma = slice.reduce((s, v) => s + v, 0) / period;
    const variance = slice.reduce((s, v) => s + Math.pow(v - sma, 2), 0) / period;
    const stdDev = Math.sqrt(variance);

    const upper = sma + stdDevMultiplier * stdDev;
    const lower = sma - stdDevMultiplier * stdDev;
    const close = ohlc[i].close;
    const percentB = (upper - lower) !== 0 ? (close - lower) / (upper - lower) : 0.5;
    const bandwidth = sma !== 0 ? ((upper - lower) / sma) * 100 : 0;

    result.push({
      date: ohlc[i].date,
      sma: Math.round(sma * 100) / 100,
      upper: Math.round(upper * 100) / 100,
      lower: Math.round(lower * 100) / 100,
      percentB: Math.round(percentB * 100) / 100,
      bandwidth: Math.round(bandwidth * 100) / 100
    });
  }
  return result;
}

/**
 * ดึงค่า Bollinger Bands ล่าสุดของหุ้นตัวเดียว พร้อมแปลผลตำแหน่งราคาเป็นภาษาคนอ่านง่าย
 */
function getBollingerBands(ticker, period) {
  period = period || 20;
  const ohlc = _getOHLCFromExternalLog(ticker);
  if (!ohlc || ohlc.length < period) {
    return { error: 'ข้อมูลราคาย้อนหลังไม่พอ (ต้องการอย่างน้อย ' + period + ' วัน มีแค่ ' + (ohlc ? ohlc.length : 0) + ' วัน)' };
  }
  const series = _calculateBollingerBands(ohlc, period, 2);
  if (!series.length) return { error: 'คำนวณ Bollinger Bands ไม่สำเร็จ' };

  const latest = series[series.length - 1];
  let position;
  if (latest.percentB >= 1) position = 'ราคาแตะ/ทะลุกรอบบน — Overbought ระวังการปรับฐาน';
  else if (latest.percentB >= 0.8) position = 'ใกล้กรอบบน';
  else if (latest.percentB <= 0) position = 'ราคาแตะ/ทะลุกรอบล่าง — Oversold อาจมีการเด้งกลับ';
  else if (latest.percentB <= 0.2) position = 'ใกล้กรอบล่าง';
  else position = 'อยู่กลางกรอบ';

  return {
    ticker, period, date: latest.date, close: ohlc[ohlc.length - 1].close,
    sma: latest.sma, upper: latest.upper, lower: latest.lower,
    percentB: latest.percentB, bandwidth: latest.bandwidth, position
  };
}

/**
 * ตรวจ Bollinger Squeeze — เทียบ Bandwidth ปัจจุบัน กับช่วงย้อนหลัง (default 126 วัน ~6 เดือน)
 * ถ้า Bandwidth อยู่ใน 10% ล่างสุดของช่วงที่ผ่านมา = แถบบีบตัวแคบผิดปกติ
 */
function getBollingerSqueezeStatus(ticker, period, lookbackDays) {
  period = period || 20;
  lookbackDays = lookbackDays || 126;

  const ohlc = _getOHLCFromExternalLog(ticker);
  if (!ohlc || ohlc.length < period + 10) {
    return { error: 'ข้อมูลไม่พอสำหรับตรวจ Squeeze' };
  }

  const series = _calculateBollingerBands(ohlc, period, 2);
  if (series.length < 30) {
    return { error: 'ข้อมูลย้อนหลังยังน้อยไป (ต้องการอย่างน้อย ~30 วันของค่า Bandwidth ถึงจะเทียบได้แม่นยำ)' };
  }

  const lookback = series.slice(-Math.min(lookbackDays, series.length));
  const bandwidths = lookback.map(s => s.bandwidth);
  const minBandwidth = Math.min(...bandwidths);
  const maxBandwidth = Math.max(...bandwidths);
  const current = series[series.length - 1];

  const range = maxBandwidth - minBandwidth;
  const percentile = range > 0 ? ((current.bandwidth - minBandwidth) / range) * 100 : 50;
  const isSqueeze = percentile <= 10;

  const status = isSqueeze
    ? '🔴 Squeeze — แถบบีบตัวแคบสุดในรอบ ' + lookback.length + ' วัน มักเป็นสัญญาณล่วงหน้าว่าจะมีการเคลื่อนไหวแรงเร็วๆ นี้'
    : percentile <= 25
      ? '🟡 แถบเริ่มแคบ — จับตาดูใกล้ๆ'
      : '🟢 แถบปกติ ไม่มีสัญญาณ Squeeze';

  return {
    ticker, bandwidth: current.bandwidth,
    minBandwidth: Math.round(minBandwidth * 100) / 100, maxBandwidth: Math.round(maxBandwidth * 100) / 100,
    percentile: Math.round(percentile * 10) / 10, isSqueeze, status, lookbackDays: lookback.length
  };
}


/**
 * Wrapper สำหรับเรียกจากหน้าเว็บ (google.script.run) — ห่อ error ให้ปลอดภัย
 */
function getBollingerBandsForStockDetail(ticker, market) {
  try {
    const r = getBollingerBands(ticker, 20);
    if (r.error) return { success: false, error: r.error };

    const squeeze = getBollingerSqueezeStatus(ticker, 20, 126);

    return {
      success: true, close: r.close, sma: r.sma, upper: r.upper, lower: r.lower,
      percentB: r.percentB, bandwidth: r.bandwidth, position: r.position,
      squeeze: squeeze.error ? null : squeeze
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * สแกน Holdings + Watchlist ทุกเช้า แจ้งเฉพาะตัวที่มีสัญญาณเด่น
 * (คะแนนสูง/ต่ำผิดปกติ หรือ Bollinger Squeeze) ไม่ส่งข้อมูลทุกตัวทุกวัน กันสแปม
 */
function dailySignalScan() {
  try {
    const data = getScreenerData();
    if (!data.success) {
      sendTelegramError('dailySignalScan', new Error(data.error));
      return;
    }

    const bullish = data.results.filter(r => !r.error && r.score >= 70);
    const bearish = data.results.filter(r => !r.error && r.score <= 30);

    const squeezed = [];
    data.results.forEach(r => {
      if (r.error) return;
      try {
        const sq = getBollingerSqueezeStatus(r.ticker, 20, 126);
        if (!sq.error && sq.isSqueeze) squeezed.push({ ticker: r.ticker, percentile: sq.percentile });
      } catch (e) {}
    });

    if (bullish.length === 0 && bearish.length === 0 && squeezed.length === 0) {
      sendTelegramSafe('📊 สรุปสัญญาณรายวัน\n\nวันนี้ไม่มีหุ้นตัวไหนมีสัญญาณเด่นเป็นพิเศษ');
      return;
    }

    let msg = '📊 สรุปสัญญาณรายวัน\n\n';

    if (bullish.length > 0) {
      msg += '🟢 สัญญาณขาขึ้นแข็งแรง (คะแนน ≥70):\n';
      bullish.forEach(r => { msg += `  • ${r.ticker} — ${r.score}/100 (${r.bullishCount}/${r.total})\n`; });
      msg += '\n';
    }
    if (bearish.length > 0) {
      msg += '🔴 สัญญาณขาลงแข็งแรง (คะแนน ≤30):\n';
      bearish.forEach(r => { msg += `  • ${r.ticker} — ${r.score}/100 (${r.bullishCount}/${r.total})\n`; });
      msg += '\n';
    }
    if (squeezed.length > 0) {
      msg += '🎯 Bollinger Squeeze (แถบบีบแคบ เตรียมเบรก):\n';
      squeezed.forEach(s => { msg += `  • ${s.ticker} — Percentile ${s.percentile}%\n`; });
      msg += '\n';
    }

    msg += '(สแกนจาก Holdings + Watchlist ทั้งหมด ' + data.count + ' ตัว)';
    sendTelegramSafe(msg);
  } catch (e) {
    logError('dailySignalScan', e);
    sendTelegramError('dailySignalScan', e);
  }
}

/**
 * รันฟังก์ชันนี้ "แค่ครั้งเดียว" ด้วยตัวเอง เพื่อตั้งเวลาสแกนอัตโนมัติทุกเช้า 08:00
 */
function setupDailySignalScanTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailySignalScan') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('dailySignalScan')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('ตั้งเวลาสแกนสัญญาณรายวันอัตโนมัติ ตอน 08:00 เรียบร้อยแล้ว');
}

function cmdScanNow(text) {
  sendTelegramSafe('⏳ กำลังสแกน Holdings + Watchlist ทั้งหมด รอสักครู่...');
  dailySignalScan();
}


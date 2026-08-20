// ============================================================
// webapp_03b_signal_overview.gs — หน้าภาพรวมสัญญาณ
// รวม Fast Signal + Portfolio Analysis ของทุกตัวใน StockMode เป็น list เดียว
// ให้หน้า "ภาพรวมสัญญาณ" ใช้แสดงการ์ด
//
// ⚠️ ต้องมี: webapp_00b_helpers.gs (getSheet, logError)
//           webapp_05_settings.gs (getStockModeMap, getEffectiveRiskParams)
//           webapp_03_analyze.gs (getFastSignal, getStockAnalysis — ฟังก์ชันเดิม)
//           webapp_02_holdings.gs (getHoldingsData — เช็คสถานะถือ/ไม่ถือ)
// ============================================================

// ══════════════════════════════════════════════════════════
// ดึงสัญญาณของทุกหุ้นใน StockMode มารวมเป็น list เดียว
// ใช้ Fast Signal สำหรับหุ้นโหมด Fast, ใช้ Portfolio Analysis สำหรับที่เหลือ
// คืนค่าเป็น array พร้อมเรียงตามความเร่งด่วน (SELL ที่หลุด stop ขึ้นก่อนเสมอ)
// ══════════════════════════════════════════════════════════
function getSignalOverviewData() {
  try {
    const modeMap = getStockModeMap();
    const tickers = Object.keys(modeMap);
    if (!tickers.length) return { success: true, items: [] };
    
    _prefetchYahooHistoryBatch(tickers, modeMap);   // ← เพิ่มบรรทัดนี้

    // ── เช็คว่าตัวไหนถืออยู่แล้ว (สำหรับ badge "ถืออยู่ X%") ──
    const heldMap = {}; // { TICKER: { unrealizedPct, valueNow } }
    try {
      const h = getHoldingsData();
      (h.us || []).forEach(x => heldMap[x.ticker] = x);
      (h.th || []).forEach(x => heldMap[x.ticker] = x);
    } catch (e) { /* ไม่มี holdings ก็ข้าม ไม่ให้ทั้งฟังก์ชันพัง */ }

    const items = [];
    tickers.forEach((ticker, idx) => {
      const cfg = modeMap[ticker];
      const market = (cfg.market === 'ไทย') ? 'TH' : 'US';
      const isFast = cfg.mode === 'Fast';

      const t0 = new Date().getTime(); // ── วัดเวลาต่อ ticker ไว้ดูใน Execution log ──
      let sig;
      try {
        sig = isFast ? getFastSignal(ticker) : getStockAnalysis(ticker);
      } catch (e) {
        logError('getSignalOverviewData:' + ticker, e);
        items.push(_buildErrorItem(ticker, market, cfg, e.message)); // ── โชว์เป็นการ์ด error แทนหายไปเงียบๆ ──
        return;
      }
      const elapsedSec = ((new Date().getTime() - t0) / 1000).toFixed(1);
      Logger.log('[' + (idx + 1) + '/' + tickers.length + '] ' + ticker + ' (' + (isFast ? 'Fast' : 'Portfolio') + ') ใช้เวลา ' + elapsedSec + ' วิ');

      if (!sig || !sig.success) {
        const errMsg = (sig && sig.error) ? sig.error : 'วิเคราะห์ไม่สำเร็จ (ไม่ทราบสาเหตุ)';
        logError('getSignalOverviewData:' + ticker + ' (sig.success=false)', new Error(errMsg));
        items.push(_buildErrorItem(ticker, market, cfg, errMsg)); // ── โชว์เป็นการ์ด error แทนหายไปเงียบๆ ──
        return;
      }

      const normalized = _normalizeSignal(sig, isFast, ticker, market, cfg, heldMap[ticker]);
      if (normalized) items.push(normalized);

      // ── กันยิง Yahoo Finance ติดกันถี่เกินไปจนโดนบล็อกชั่วคราว (เหมือนที่ทำใน updateWatchlistPricesWeb) ──
      // ลบ if (idx < tickers.length - 1) ;
    });

    // เรียงความเร่งด่วน: SELL ที่หลุด stop → SELL อื่นๆ → BUY → WATCH
    const urgencyOrder = { sell_urgent: 0, sell: 1, buy: 2, watch: 3, error: 4 };
    items.sort((a, b) => (urgencyOrder[a.sortKey] ?? 9) - (urgencyOrder[b.sortKey] ?? 9));

    return { success: true, items };
  } catch (e) {
    logError('getSignalOverviewData', e);
    return { success: false, error: e.message, items: [] };
  }
}


function _prefetchYahooHistoryBatch(tickers, modeMap) {
  const cache = CacheService.getScriptCache();
  const requests = [], meta = [];

  tickers.forEach(ticker => {
    const cfg = modeMap[ticker];
    const market = (cfg.market === 'ไทย') ? 'TH' : 'US';
    const symbol = (market === 'TH') ? (ticker.toUpperCase() + '.BK') : ticker.toUpperCase();
    const range = (cfg.mode === 'Fast') ? '3mo' : '1y';
    const cacheKey = 'yh_' + symbol + '_' + range;
    if (cache.get(cacheKey)) return;
    requests.push({
      url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?interval=1d&range=' + range,
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    meta.push({ cacheKey });
  });

  if (!requests.length) return;

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests); // ← ยิงขนานทีเดียวทั้งชุด
  } catch (e) {
    logError('_prefetchYahooHistoryBatch:fetchAll', e);
    return; // ไม่มี cache ก็ไม่พัง — ฟังก์ชันเดิม fallback ไปยิงเองทีละตัวตามปกติ
  }

  responses.forEach((res, i) => {
    try {
      if (res.getResponseCode() === 200) cache.put(meta[i].cacheKey, res.getContentText(), 60);
    } catch (e) {}
  });
}


// ── แปลงผลลัพธ์จาก getFastSignal()/getStockAnalysis() ให้เป็นรูปแบบเดียวกัน
//    สำหรับการ์ดภาพรวม (ทั้งสองฟังก์ชันคืนโครงสร้างต่างกัน ต้อง normalize) ──
function _normalizeSignal(sig, isFast, ticker, market, cfg, heldInfo) {
  let signalType, confidence, reasons, warnings, urgentSell, ma20;

  if (isFast) {
    // getFastSignal() เดิม: มี decision, decisionClass, reasonsFor, reasonsAgainst, belowStop
    signalType = _mapDecisionToSignal(sig.decisionClass);
    confidence = (sig.reasonsFor || []).length; // ใช้จำนวนเหตุผลสนับสนุนเป็น proxy ความมั่นใจ (0-4)
    reasons = sig.reasonsFor || [];
    warnings = sig.reasonsAgainst || [];
    urgentSell = !!sig.belowStop;
    ma20 = null; // getFastSignal ปัจจุบันไม่คืนค่า MA มาตรงๆ (ใช้ EMA5/EMA20 ภายใน) — ปล่อย null ไว้ก่อน
  } else {
    // getStockAnalysis() เดิม: มี decision, decClass, reasons, warnings, ma20
    signalType = _mapDecisionToSignal(sig.decClass);
    confidence = (sig.reasons || []).length;
    reasons = sig.reasons || [];
    warnings = sig.warnings || [];
    urgentSell = sig.decClass === 'stop' && sig.plClass === 'stop';
    ma20 = sig.ma20 || null;
  }

  // ── ระยะห่างจาก MA20 (%) แทนที่ "Upside" เดิม ──
  // เหตุผลที่ตัด Upside ออก: ระบบยังไม่มี valuation model (P/E เป้าหมาย/DCF) ที่คำนวณราคาเป้าหมายได้จริง
  // การใส่ตัวเลข "Upside" แบบคาดเดาจาก MA/ATR อาจทำให้เข้าใจผิดว่าเป็นการพยากรณ์ราคา จึงใช้ค่าที่วัดได้จริงแทน
  const distFromMA20Pct = (ma20 && sig.price) ? ((sig.price - ma20) / ma20) * 100 : null;

  const sortKey = signalType === 'sell' ? (urgentSell ? 'sell_urgent' : 'sell') : signalType;

  return {
    ticker, market,
    companyName: _getCompanyNameSafe(ticker), // fallback เป็น ticker ถ้าไม่มีฟังก์ชัน/ข้อมูลชื่อบริษัท
    tradeStyle: cfg.trendGroup || '',
    assetType: cfg.assetType || '',
    mode: cfg.mode || 'Portfolio',
    signal: signalType,          // 'buy' | 'watch' | 'sell'
    confidence: Math.min(4, confidence), // 0-4 ใช้ทำ confidence dots
    price: sig.price,
    distFromMA20Pct,             // แทนที่ upside — null ถ้าไม่มี MA20 ให้เทียบ (เช่น Fast mode)
    reasons: reasons.slice(0, 4),
    warnings: warnings.slice(0, 4),
    urgentSell,
    held: !!heldInfo,
    heldPct: heldInfo ? (parseFloat(heldInfo.unrealizedPct) || 0) * 100 : null,
    sortKey,
    updatedAt: sig.updatedAt || ''
  };
}

// ── ดึงชื่อเต็มบริษัท ถ้ามีฟังก์ชัน getCompanyName() อยู่ใน stockinfo.gs (ไฟล์เดิม) ให้ใช้เลย
//    ถ้าไม่มี/พัง fallback เป็น ticker เฉยๆ (ไม่ใช้ cfg.note เพราะเป็นช่องหมายเหตุอิสระ ไม่ใช่ชื่อบริษัท) ──
function _getCompanyNameSafe(ticker) {
  try {
    if (typeof getCompanyName === 'function') {
      const name = getCompanyName(ticker);
      if (name) return name;
    }
  } catch (e) { /* เงียบไว้ — ไม่ให้กระทบทั้งลิสต์ */ }
  return ticker;
}

// ── map decClass/decisionClass เดิม (safe/warn/stop) → signal (buy/watch/sell) ──
function _mapDecisionToSignal(cls) {
  if (cls === 'safe') return 'buy';
  if (cls === 'stop') return 'sell';
  return 'watch';
}

// ── สร้าง item แบบ "วิเคราะห์ไม่สำเร็จ" — โชว์เป็นการ์ดในลิสต์แทนหายไปเงียบๆ
//    ให้ผู้ใช้เห็นสาเหตุตรงๆ ในหน้าเว็บ ไม่ต้องเปิด Execution log ทุกครั้ง ──
function _buildErrorItem(ticker, market, cfg, errorMessage) {
  return {
    ticker, market,
    companyName: _getCompanyNameSafe(ticker),
    tradeStyle: cfg.trendGroup || '',
    assetType: cfg.assetType || '',
    mode: cfg.mode || 'Portfolio',
    signal: 'error',
    confidence: 0,
    price: null,
    distFromMA20Pct: null,
    reasons: [],
    warnings: [],
    errorMessage,
    urgentSell: false,
    held: false,
    heldPct: null,
    sortKey: 'error',
    updatedAt: ''
  };
}

// ══════════════════════════════════════════════════════════
// นับจำนวนสัญญาณแต่ละแบบ (ใช้ทำแถบสรุป BUY/WATCH/SELL บนหน้าเว็บ)
// เรียกแยกจาก getSignalOverviewData() เพื่อให้ frontend cache list ไว้ใช้ซ้ำได้
// โดยไม่ต้องยิง request ใหม่แค่นับจำนวน
// ══════════════════════════════════════════════════════════
function countSignalTypes(items) {
  const counts = { buy: 0, watch: 0, sell: 0 };
  (items || []).forEach(x => { if (counts[x.signal] !== undefined) counts[x.signal]++; });
  return counts;
}



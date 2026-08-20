// ============================================================
// webapp_07_watchlist.gs — หน้า Watchlist
// เพิ่ม/ลบ(soft)/แก้โน้ต หุ้นที่ติดตาม + ดึงราคาสด + วิเคราะห์เทคนิคอล
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
//
// ⚠️ ต้องมี webapp_00b_helpers.gs (getSheet, logError) อยู่ในโปรเจกต์แล้ว
// ⚠️ ใช้ _getNextEmptyRow() ตัวเดียวกับ webapp_00_main.gs (ไม่ประกาศซ้ำ)
//
// โครงสร้างชีต "⭐ Watchlist" (header แถว 6, ข้อมูลเริ่มแถว 7):
//   A=Ticker  B=Market(US/TH)  C=tradeStyle  D=targetPrice  E=supportPrice
//   F=note    G=วันที่ Watchlist  H=ราคา ณ วันที่เพิ่ม  I=ราคา ณ ปัจจุบัน
//   J=เปลี่ยนแปลง %  K=สถานะ (watchlist / cancel)
// ============================================================

const WATCHLIST_SHEET = {
  NAME: '⭐ Watchlist',   // ⚠️ ต้องตรงกับชื่อแท็บจริงเป๊ะๆ (รวม emoji/เว้นวรรค) — เช็คถ้า error "ไม่พบชีต"
  START_ROW: 7
};

// ══════════════════════════════════════════════════════════
// ดึงรายการ Watchlist ทั้งหมดที่ยัง active (สถานะ = watchlist)
// คำนวณ diffFromTargetPct + zone (ready/watch/far) แบบสด ไม่เก็บลงชีต
// เพิ่ม: recommendation (ข้อความสั้นจาก zone) + lowestBuyPrice (ถ้าเคยซื้อมาก่อน)
// ══════════════════════════════════════════════════════════
function getWatchlistData() {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < WATCHLIST_SHEET.START_ROW) return { success: true, items: [] };

    const numRows = lastRow - WATCHLIST_SHEET.START_ROW + 1;
    const rows = sheet.getRange(WATCHLIST_SHEET.START_ROW, 1, numRows, 11).getValues();

    // ── คำนวณราคาซื้อต่ำสุดต่อ ticker ล่วงหน้าครั้งเดียว (แยก US/TH) กันดึงชีตซ้ำทุกแถว ──
    const lowestBuyMaps = _getLowestBuyPriceMapsByMarket();

    const items = [];
    rows.forEach((row, i) => {
      const ticker = String(row[0] || '').trim().toUpperCase();
      const status = String(row[10] || 'watchlist').trim().toLowerCase();
      if (!ticker || status === 'cancel') return;

      const market       = String(row[1] || 'US').trim().toUpperCase();
      const targetPrice  = parseFloat(row[3]) || 0;
      const supportPrice = row[4] !== '' ? parseFloat(row[4]) : null;
      const currentPrice = parseFloat(row[8]) || parseFloat(row[7]) || 0;

      const diffFromTargetPct = targetPrice > 0 ? ((currentPrice - targetPrice) / targetPrice) * 100 : null;
      let zone = 'far';
      if (diffFromTargetPct !== null) {
        if (diffFromTargetPct <= 0) zone = 'ready';
        else if (diffFromTargetPct <= 10) zone = 'watch';
      }

      // ── ราคาต่ำสุดที่เคยซื้อ — ถ้าไม่เคยซื้อมาก่อนเลยจะเป็น null (ไม่แสดงในการ์ด) ──
      const lowestMap = (market === 'TH') ? lowestBuyMaps.th : lowestBuyMaps.us;
      const lowestBuyPrice = lowestMap[ticker] || null;

      items.push({
        rowIndex: WATCHLIST_SHEET.START_ROW + i,
        ticker, market,
        tradeStyle: row[2] || '',
        targetPrice, supportPrice,
        note: row[5] || '',
        dateAdded: row[6] ? Utilities.formatDate(new Date(row[6]), 'Asia/Bangkok', 'dd/MM/yyyy') : '',
        priceAdded: parseFloat(row[7]) || null,
        currentPrice,
        changePct: parseFloat(row[9]) || 0,
        diffFromTargetPct, zone,
        recommendation: _zoneToRecommendation(zone),
        lowestBuyPrice
      });
    });

    // เรียง ready ก่อน แล้ว watch แล้ว far
    const zoneOrder = { ready: 0, watch: 1, far: 2 };
    items.sort((a, b) => zoneOrder[a.zone] - zoneOrder[b.zone]);

    return { success: true, items };
  } catch (e) {
    logError('getWatchlistData', e);
    return { success: false, error: e.message, items: [] };
  }
}

// ── แปลง zone เป็นข้อความคำแนะนำสั้นๆ ที่โชว์บนการ์ด
//    ใช้ zone ที่คำนวณจาก targetPrice อยู่แล้ว ไม่เรียกวิเคราะห์เทคนิคอลเพิ่ม
//    (กันหน้าโหลดช้าถ้ามีหุ้นในลิสต์เยอะ — วิเคราะห์เต็มรูปแบบทำตอนกด "วิเคราะห์" แทน) ──
function _zoneToRecommendation(zone) {
  if (zone === 'ready') return 'พร้อมซื้อ';
  if (zone === 'watch') return 'รอสัญญาณเพิ่ม';
  return 'เฝ้าดู';
}

// ── สร้าง map ราคาซื้อต่ำสุดต่อ ticker แยกตามตลาด (US/TH) จาก Transaction Log จริง
//    reuse ฟังก์ชัน getLowestBuyPriceMap() ที่มีอยู่แล้วใน webapp_02_holdings.gs
//    คำนวณครั้งเดียวต่อการเรียก getWatchlistData() 1 ครั้ง ไม่ใช่ต่อรายการ (กันดึงชีตซ้ำ) ──
function _getLowestBuyPriceMapsByMarket() {
  let usMap = {}, thMap = {};
  try {
    const usLogRows = getSheet(SHEETS.US_TRANS).getDataRange().getValues();
    usMap = getLowestBuyPriceMap(usLogRows);
  } catch (e) { logError('_getLowestBuyPriceMapsByMarket:US', e); }
  try {
    const thLogRows = getSheet(SHEETS.TH_TRANS).getDataRange().getValues();
    thMap = getLowestBuyPriceMap(thLogRows);
  } catch (e) { logError('_getLowestBuyPriceMapsByMarket:TH', e); }
  return { us: usMap, th: thMap };
}



// เพิ่มฟังก์ชัน backend เบาๆ ที่ดึง Yahoo Finance ของ**ทุกตัวพร้อมกันในคำขอเดียว** (`UrlFetchApp.fetchAll`) แทนยิงทีละตัว และดึงย้อนหลังแค่ 3 เดือน (พอสำหรับ MA20/50) แทน 1 ปี

function getWatchlistTrendsBatch(items) {
  const result = {};
  if (!items || !items.length) return result;

  const requests = items.map(it => {
    const symbol = (String(it.market).toUpperCase() === 'TH')
      ? (String(it.ticker).toUpperCase() + '.BK') : String(it.ticker).toUpperCase();
    return {
      url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?range=3mo&interval=1d',
      muteHttpExceptions: true
    };
  });

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests); // ← ยิงขนานทีเดียวทั้งชุด แทนที่จะทีละตัว
  } catch (e) {
    logError('getWatchlistTrendsBatch:fetchAll', e);
    items.forEach(it => { result[it.ticker + '_' + it.market] = { success: false }; });
    return result;
  }

  items.forEach((it, i) => {
    const key = it.ticker + '_' + it.market;
    try {
      const res = responses[i];
      if (res.getResponseCode() !== 200) { result[key] = { success: false }; return; }
      const json = JSON.parse(res.getContentText());
      const chartResult = json.chart && json.chart.result && json.chart.result[0];
      const closesRaw = chartResult && chartResult.indicators && chartResult.indicators.quote[0]
        ? chartResult.indicators.quote[0].close : null;
      if (!closesRaw) { result[key] = { success: false }; return; }
      const closes = closesRaw.filter(v => v !== null && v !== undefined);
      if (closes.length < 20) { result[key] = { success: false }; return; }

      const ma20 = _wlCalcSMA(closes, 20);
      const ma50 = _wlCalcSMA(closes, Math.min(50, closes.length));
      const price = closes[closes.length - 1];

      let trendClass;
      if (ma20 && ma50 && price > ma20 && ma20 > ma50) trendClass = 'safe';
      else if (ma20 && ma50 && price < ma20 && ma20 < ma50) trendClass = 'stop';
      else trendClass = 'warn';

      result[key] = { success: true, trendClass };
    } catch (e) {
      result[key] = { success: false };
    }
  });

  return result;
}

// ══════════════════════════════════════════════════════════
// เพิ่มหุ้นเข้า Watchlist — เช็คซ้ำ + ดึงราคาปัจจุบันจาก Yahoo Finance มา snapshot
// ══════════════════════════════════════════════════════════
function addWatchlistItem(data) {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const ticker = String(data.ticker || '').trim().toUpperCase();
    const market = (data.market === 'TH') ? 'TH' : 'US';
    if (!ticker) return { success: false, error: 'กรุณาระบุ ticker' };
    if (!data.targetPrice || parseFloat(data.targetPrice) <= 0) {
      return { success: false, error: 'กรุณาระบุราคาเป้าหมายที่ถูกต้อง' };
    }

    // กันเพิ่มซ้ำ ticker เดิมที่ยัง active อยู่ (ตลาดเดียวกัน)
    const existing = getWatchlistData();
    const dup = (existing.items || []).some(x => x.ticker === ticker && x.market === market);
    if (dup) return { success: false, error: ticker + ' มีใน Watchlist อยู่แล้ว' };

    const quote = _wlFetchYahooQuote(ticker, market);
    if (!quote) return { success: false, error: 'ไม่พบหุ้น "' + ticker + '" ตรวจสอบชื่อ ticker และตลาดอีกครั้ง' };

    const row = _getNextEmptyRow(sheet, 1, WATCHLIST_SHEET.START_ROW); // เช็คจาก col A (Ticker)
    const now = new Date();

    sheet.getRange(row, 1).setValue(ticker);
    sheet.getRange(row, 2).setValue(market);
    sheet.getRange(row, 3).setValue(data.tradeStyle || '');
    sheet.getRange(row, 4).setValue(parseFloat(data.targetPrice));
    sheet.getRange(row, 5).setValue(data.supportPrice ? parseFloat(data.supportPrice) : '');
    sheet.getRange(row, 6).setValue(data.note || '');
    sheet.getRange(row, 7).setValue(now).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 8).setValue(quote.price);
    sheet.getRange(row, 9).setValue(quote.price);
    sheet.getRange(row, 10).setValue(0);
    sheet.getRange(row, 11).setValue('watchlist');

    sheet.getRange(row, 4, 1, 3).setNumberFormat('#,##0.00');
    sheet.getRange(row, 8, 1, 2).setNumberFormat('#,##0.00');
    sheet.getRange(row, 10).setNumberFormat('0.00"%"');

    return { success: true, row, price: quote.price, longName: quote.longName };
  } catch (e) {
    logError('addWatchlistItem', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// นำออกจาก Watchlist — soft delete (set สถานะ = cancel) กันข้อมูล/โน้ตหาย
// ══════════════════════════════════════════════════════════
function removeWatchlistItem(ticker, market) {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const rowIdx = _findWatchlistRow(sheet, ticker, market);
    if (rowIdx === -1) return { success: false, error: 'ไม่พบ ' + ticker + ' ใน Watchlist' };

    sheet.getRange(rowIdx, 11).setValue('cancel');
    return { success: true, ticker };
  } catch (e) {
    logError('removeWatchlistItem', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// แก้ไขโน้ต/แผนของหุ้นใน Watchlist
// ══════════════════════════════════════════════════════════
function updateWatchlistNote(ticker, market, note) {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const rowIdx = _findWatchlistRow(sheet, ticker, market);
    if (rowIdx === -1) return { success: false, error: 'ไม่พบ ' + ticker + ' ใน Watchlist' };

    sheet.getRange(rowIdx, 6).setValue(note || '');
    return { success: true };
  } catch (e) {
    logError('updateWatchlistNote', e);
    return { success: false, error: e.message };
  }
}

// ── หา row ของ ticker ที่ยัง active (สถานะ = watchlist) — ใช้ร่วมกันหลายฟังก์ชัน ──
function _findWatchlistRow(sheet, ticker, market) {
  const lastRow = sheet.getLastRow();
  if (lastRow < WATCHLIST_SHEET.START_ROW) return -1;

  const numRows = lastRow - WATCHLIST_SHEET.START_ROW + 1;
  const rows = sheet.getRange(WATCHLIST_SHEET.START_ROW, 1, numRows, 11).getValues();
  const tUpper = String(ticker || '').trim().toUpperCase();

  for (let i = 0; i < rows.length; i++) {
    const rTicker = String(rows[i][0] || '').trim().toUpperCase();
    const rMarket = String(rows[i][1] || '').trim().toUpperCase();
    const rStatus = String(rows[i][10] || 'watchlist').trim().toLowerCase();
    if (rTicker === tUpper && (!market || rMarket === market) && rStatus === 'watchlist') {
      return WATCHLIST_SHEET.START_ROW + i;
    }
  }
  return -1;
}

// ══════════════════════════════════════════════════════════
// อัปเดตราคาปัจจุบัน + เปลี่ยนแปลง % ของทุกตัวใน Watchlist
// เรียกจากปุ่ม "รีเฟรช" ในหน้า Watchlist (แยกจากปุ่มอัปเดตราคา Holdings เดิม)
// ══════════════════════════════════════════════════════════
function updateWatchlistPricesWeb() {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < WATCHLIST_SHEET.START_ROW) return { success: true, updated: 0, failed: [] };

    const numRows = lastRow - WATCHLIST_SHEET.START_ROW + 1;
    const rows = sheet.getRange(WATCHLIST_SHEET.START_ROW, 1, numRows, 11).getValues();

    let updated = 0;
    const failed = []; // ← เพิ่ม

    rows.forEach((row, i) => {
      const ticker = String(row[0] || '').trim();
      const status = String(row[10] || 'watchlist').trim().toLowerCase();
      if (!ticker || status === 'cancel') return;

      const market = String(row[1] || 'US').trim();
      const priceAdded = parseFloat(row[7]) || 0;

      const quote = _wlFetchYahooQuote(ticker, market);
      if (!quote) { failed.push(ticker); return; } // ← เปลี่ยนจาก return เฉยๆ

      const r = WATCHLIST_SHEET.START_ROW + i;
      sheet.getRange(r, 9).setValue(quote.price);
      if (priceAdded > 0) {
        sheet.getRange(r, 10).setValue(((quote.price - priceAdded) / priceAdded) * 100);
      }
      updated++;
      Utilities.sleep(200);
    });

    return {
      success: true, updated, failed, // ← ส่ง failed กลับไปด้วย
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('updateWatchlistPricesWeb', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// ค้นหา ticker เพื่อเพิ่มเข้า Watchlist
// รวมลิสต์จาก getPortfolioData() เดิม (ATR list) + tag ว่าตัวไหนถืออยู่แล้ว/
// อยู่ใน Watchlist แล้ว (กันเพิ่มซ้ำตั้งแต่ตอนค้นหา)
// ══════════════════════════════════════════════════════════
function getWatchlistSearchList() {
  try {
    const portfolio = getPortfolioData(); // ฟังก์ชันเดิมจาก webapp_03_analyze.gs

    const heldTickers = new Set();
    try {
      const h = getHoldingsData();
      (h.us || []).forEach(x => heldTickers.add(x.ticker));
      (h.th || []).forEach(x => heldTickers.add(x.ticker));
    } catch (e) { /* ไม่มี holdings ก็ข้ามไป ไม่ให้ทั้งฟังก์ชันพัง */ }

    const watchedKeys = new Set();
    try {
      const w = getWatchlistData();
      (w.items || []).forEach(x => watchedKeys.add(x.ticker + '|' + x.market));
    } catch (e) { /* เช่นกัน */ }

    function tag(list, market) {
      return (list || []).map(r => ({
        ...r,
        alreadyHeld: heldTickers.has(String(r.symbol || '').toUpperCase()),
        alreadyWatched: watchedKeys.has(String(r.symbol || '').toUpperCase() + '|' + market)
      }));
    }

    return { success: true, us: tag(portfolio.us, 'US'), th: tag(portfolio.th, 'TH') };
  } catch (e) {
    logError('getWatchlistSearchList', e);
    return { success: false, error: e.message, us: [], th: [] };
  }
}

// ══════════════════════════════════════════════════════════
// เช็ก ticker เดี่ยวๆ ตอนพิมพ์เอง (โหมด "พิมพ์เอง") — ใช้ validate ก่อนบันทึก
// ══════════════════════════════════════════════════════════
function lookupTickerPrice(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    if (!ticker) return { success: false, error: 'กรุณาระบุ ticker' };

    const quote = _wlFetchYahooQuote(ticker, market);
    if (!quote) return { success: false, error: 'ไม่พบหุ้น "' + ticker + '" ตรวจสอบชื่อ ticker อีกครั้ง' };

    return { success: true, ticker, price: quote.price, currency: quote.currency, longName: quote.longName };
  } catch (e) {
    logError('lookupTickerPrice', e);
    return { success: false, error: 'เชื่อมต่อ Yahoo Finance ไม่สำเร็จ ลองใหม่อีกครั้ง' };
  }
}

// ══════════════════════════════════════════════════════════
// วิเคราะห์เทคนิคอลสำหรับหุ้นใน Watchlist (ไม่มีตำแหน่งถือ — ไม่มี P/L)
// เทียบกับ targetPrice/supportPrice ที่ผู้ใช้ตั้งไว้ แทนการเทียบ entry price
// ══════════════════════════════════════════════════════════
function getWatchlistAnalysis(ticker, market, targetPrice, supportPrice) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const hist = _wlFetchYahooHistory(ticker, market, '1y', '1d');
    if (!hist || !hist.closes || hist.closes.length < 20) {
      return { success: false, error: 'ไม่พบข้อมูลราคาย้อนหลังของ ' + ticker + ' เพียงพอสำหรับวิเคราะห์' };
    }

    const closes = hist.closes;
    const volumes = hist.volumes;
    const price = closes[closes.length - 1];

    const ma20  = _wlCalcSMA(closes, 20);
    const ma50  = _wlCalcSMA(closes, 50);
    const ma200 = _wlCalcSMA(closes, 200);
    const rsi   = _wlCalcRSI(closes, 14);
    const volNow = volumes.length ? volumes[volumes.length - 1] : null;
    const volAvg = _wlCalcSMA(volumes, 20);
     const pivotLevels = _wlCalcPivotLevels(hist.highs, hist.lows, closes);
    const pivotTakeaways = buildPivotKeyTakeaways(pivotLevels, price);



    // ── Trend ──
    let trendSignal, trendClass;
    if (ma20 && ma50 && price > ma20 && ma20 > ma50) {
      trendSignal = 'ขาขึ้นชัดเจน — ราคาอยู่เหนือ MA20/50'; trendClass = 'safe';
    } else if (ma20 && ma50 && price < ma20 && ma20 < ma50) {
      trendSignal = 'ขาลงชัดเจน — ราคาต่ำกว่า MA20/50'; trendClass = 'stop';
    } else {
      trendSignal = 'แนวโน้มไม่ชัดเจน / Sideway'; trendClass = 'warn';
    }

    // ── RSI ──
    let rsiSignal, rsiClass;
    if (rsi === null) {
      rsiSignal = 'ไม่มีข้อมูลเพียงพอคำนวณ RSI'; rsiClass = 'warn';
    } else if (rsi >= 70) {
      rsiSignal = 'RSI ' + rsi.toFixed(1) + ' — Overbought ระวังราคาย่อ'; rsiClass = 'warn';
    } else if (rsi <= 30) {
      rsiSignal = 'RSI ' + rsi.toFixed(1) + ' — Oversold มีโอกาสเด้งกลับ'; rsiClass = 'safe';
    } else {
      rsiSignal = 'RSI ' + rsi.toFixed(1) + ' — อยู่ในโซนกลาง'; rsiClass = 'safe';
    }

    // ── Volume ──
    let volSignal, volClass;
    if (volAvg && volNow && volNow > volAvg * 1.3) {
      volSignal = 'Volume สูงกว่าเฉลี่ยชัดเจน — มีแรงซื้อ/ขายเข้ามา'; volClass = 'safe';
    } else {
      volSignal = 'Volume ปกติ ไม่มีแรงซื้อ/ขายชัดเจน'; volClass = 'warn';
    }

    // ── ระยะห่างจากแผน ──
    const tp = parseFloat(targetPrice) || 0;
    const sp = supportPrice ? parseFloat(supportPrice) : null;
    const diffFromTargetPct = tp > 0 ? ((price - tp) / tp) * 100 : null;
    const nearSupport = sp ? (price <= sp * 1.03) : false;

    // ── คำตัดสินรวม ──
    let decision, decClass;
    if (diffFromTargetPct !== null && diffFromTargetPct <= 0 && trendClass !== 'stop' && (rsi === null || rsi < 70)) {
      decision = 'เข้าซื้อได้'; decClass = 'safe';
    } else if ((diffFromTargetPct !== null && diffFromTargetPct <= 10) || nearSupport) {
      decision = 'รอสัญญาณเพิ่มเติม'; decClass = 'warn';
    } else {
      decision = 'ยังไม่เหมาะ ราคายังไกลเป้าหมาย'; decClass = 'stop';
    }

    // ── เหตุผลประกอบ ──
    const reasons = [];
    const warnings = [];
    if (diffFromTargetPct !== null) {
      if (diffFromTargetPct <= 0) {
        reasons.push('ราคาปัจจุบันต่ำกว่า/เท่าเป้าหมายแล้ว (' + Math.abs(diffFromTargetPct).toFixed(1) + '% ต่ำกว่าเป้าหมาย)');
      } else {
        warnings.push('ราคายังสูงกว่าเป้าหมาย ' + diffFromTargetPct.toFixed(1) + '%');
      }
    }
    if (trendClass === 'stop') warnings.push('เทรนด์เป็นขาลง ต่ำกว่า MA20/50 — ควรรอสัญญาณกลับตัวก่อน');
    else if (trendClass === 'safe') reasons.push('เทรนด์เป็นขาขึ้น ราคายืนเหนือ MA หลัก');
    if (rsi !== null && rsi >= 70) warnings.push('RSI สูง เสี่ยง overbought หากซื้อตอนนี้');
        if (nearSupport) reasons.push('ราคาปัจจุบันใกล้แนวรับที่ตั้งไว้ (' + sp + ')');
    reasons.push(...pivotTakeaways.reasons);
    warnings.push(...pivotTakeaways.warnings);


    return {
      success: true, ticker, market,
      cur: market === 'TH' ? '฿' : '$',
      isTH: market === 'TH',
      price, targetPrice: tp, supportPrice: sp,
      diffFromTargetPct, nearSupport,
      ma20, ma50, ma200, rsi, volNow, volAvg,
      pivot: pivotLevels,   // ← เพิ่มบรรทัดนี้
      trendSignal, trendClass, rsiSignal, rsiClass, volSignal, volClass,
      decision, decClass, reasons, warnings,
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('getWatchlistAnalysis', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// Yahoo Finance helpers — ใช้ร่วมกันทั้งไฟล์นี้
// (แยกจาก updateThaiPrice()/updateUSPrice() เดิมใน update.gs เพราะที่นี่
//  ต้องดึงแบบ "เดี่ยวๆ ตามชื่อที่พิมพ์" ไม่ใช่ไล่ทั้งชีต Holdings)
// ══════════════════════════════════════════════════════════
function _wlFetchYahooQuote(ticker, market) {
  const symbol = (market === 'TH') ? (ticker.toUpperCase() + '.BK') : ticker.toUpperCase();
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol);
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" }
    });
    if (res.getResponseCode() !== 200) return null;
    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result || !result.meta || !result.meta.regularMarketPrice) return null;

    return {
      price: result.meta.regularMarketPrice,
      currency: result.meta.currency,
      longName: result.meta.longName || result.meta.shortName || ticker,
      symbol
    };
  } catch (e) {
    logError('fetchYahooQuote', e);
    return null;
  }
}


function _wlFetchYahooHistory(ticker, market, rangeStr, intervalStr) {
  rangeStr = rangeStr || '1y';
  intervalStr = intervalStr || '1d';
  const symbol = (market === 'TH') ? (ticker.toUpperCase() + '.BK') : ticker.toUpperCase();
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol)
    + '?range=' + rangeStr + '&interval=' + intervalStr;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result || !result.indicators || !result.indicators.quote || !result.indicators.quote[0]) return null;

    const closesRaw = result.indicators.quote[0].close || [];
    const volsRaw   = result.indicators.quote[0].volume || [];
    const highsRaw  = result.indicators.quote[0].high || [];   // ← เพิ่ม
    const lowsRaw   = result.indicators.quote[0].low || [];    // ← เพิ่ม
    const closes  = closesRaw.filter(v => v !== null && v !== undefined);
    const volumes = volsRaw.filter(v => v !== null && v !== undefined);
    const highs   = highsRaw.filter(v => v !== null && v !== undefined);  // ← เพิ่ม
    const lows    = lowsRaw.filter(v => v !== null && v !== undefined);   // ← เพิ่ม

    return { closes, volumes, highs, lows, meta: result.meta };

  } catch (e) {
    logError('fetchYahooHistory', e);
    return null;
  }
}

// ── ตัวชี้วัดเทคนิคอลพื้นฐาน ──
function _wlCalcSMA(arr, period) {
  if (!arr || arr.length < period) return null;
  const slice = arr.slice(-period);
  return slice.reduce((a, b) => a + b, 0) / period;
}

function _wlCalcRSI(closes, period) {
  period = period || 14;
  if (!closes || closes.length < period + 1) return null;

  let gains = 0, losses = 0;
  for (let i = closes.length - period; i < closes.length; i++) {
    const diff = closes[i] - closes[i - 1];
    if (diff >= 0) gains += diff; else losses -= diff;
  }
  const avgGain = gains / period, avgLoss = losses / period;
  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - (100 / (1 + rs));
}

// ── Pivot Point (R1-R3/S1-S3) จาก High/Low/Close ของ "วันล่าสุด" ที่ Yahoo ส่งมา ──
// สูตรเดียวกับ getPivotLevels() ในหน้าหุ้นที่ถือ (webapp_09_external_history.gs) — ให้ตรงกันทั้งแอป
function _wlCalcPivotLevels(highs, lows, closes) {
  try {
    if (!highs || !lows || !closes || !highs.length || !lows.length || !closes.length) {
      return { available: false, reason: 'ไม่มีข้อมูล High/Low ล่าสุดสำหรับคำนวณ Pivot Point' };
    }
    const high = highs[highs.length - 1];
    const low  = lows[lows.length - 1];
    const close = closes[closes.length - 1];
    if ([high, low, close].some(v => v === null || v === undefined || isNaN(v))) {
      return { available: false, reason: 'ข้อมูล High/Low/Close ของวันล่าสุดไม่ครบ — คำนวณ Pivot ไม่ได้' };
    }

    const pivot = (high + low + close) / 3;
    const r1 = 2 * pivot - low;
    const r2 = pivot + (high - low);
    const r3 = high + 2 * (pivot - low);
    const s1 = 2 * pivot - high;
    const s2 = pivot - (high - low);
    const s3 = low - 2 * (high - pivot);

    const allLevels = [
      { key: 'r3', price: r3 }, { key: 'r2', price: r2 }, { key: 'r1', price: r1 },
      { key: 's1', price: s1 }, { key: 's2', price: s2 }, { key: 's3', price: s3 }
    ];
    const nearestResistanceLevel = allLevels.filter(l => l.price > close).sort((a, b) => a.price - b.price)[0] || null;
    const nearestSupportLevel = allLevels.filter(l => l.price < close).sort((a, b) => b.price - a.price)[0] || null;

    return {
      available: true, pivot,
      resistance: { r1, r2, r3 },
      support: { s1, s2, s3 },
      nearestResistanceLevel, nearestSupportLevel
    };
  } catch (e) {
    logError('_wlCalcPivotLevels', e);
    return { available: false, reason: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// อัปเดตแผน (ราคาเป้าหมาย + แนวรับ) ของหุ้นใน Watchlist — ต่างจาก
// updateWatchlistNote() ที่แก้ได้แค่โน้ต ตัวนี้แก้ตัวเลขจริงในชีต (col D, E)
// ══════════════════════════════════════════════════════════
function updateWatchlistPlan(ticker, market, targetPrice, supportPrice) {
  try {
    const sheet = getSheet(WATCHLIST_SHEET.NAME);
    const rowIdx = _findWatchlistRow(sheet, ticker, market);
    if (rowIdx === -1) return { success: false, error: 'ไม่พบ ' + ticker + ' ใน Watchlist' };

    const tp = parseFloat(targetPrice);
    if (!tp || tp <= 0) return { success: false, error: 'กรุณาระบุราคาเป้าหมายที่ถูกต้อง' };

    sheet.getRange(rowIdx, 4).setValue(tp); // targetPrice
    if (supportPrice !== null && supportPrice !== undefined && supportPrice !== '') {
      const sp = parseFloat(supportPrice);
      if (!sp || sp <= 0) return { success: false, error: 'ราคาแนวรับไม่ถูกต้อง' };
      sheet.getRange(rowIdx, 5).setValue(sp);
    } else {
      sheet.getRange(rowIdx, 5).setValue('');
    }

    return { success: true, targetPrice: tp, supportPrice: supportPrice ? parseFloat(supportPrice) : null };
  } catch (e) {
    logError('updateWatchlistPlan', e);
    return { success: false, error: e.message };
  }
}



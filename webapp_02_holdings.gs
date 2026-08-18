// ============================================================
// webapp_02_holdings.gs — หน้า Holdings
// ยอดถือครองปัจจุบัน (US/TH/กองทุน) + ปุ่มอัปเดตราคา + Financial Goal/CAGR
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
// ============================================================

// ══════════════════════════════════════════════════════════
// Portfolio Holdings — getHoldingsData()
// ✅ ใช้ getHoldings() / getFundHoldings() ตัวเดียวกับที่ data.gs ใช้อยู่แล้ว
//    (SHEETS.US_HOLD / SHEETS.TH_HOLD, START_ROW.HOLD = row 7)
//    ไม่เขียน readHoldingsSheet ใหม่ — เลี่ยง column mapping ขัดกันแบบ ATR
//    ✅ เพิ่ม fund: กองทุนรวม (ใช้ getFundHoldings() เดิมจาก data.gs)
// ══════════════════════════════════════════════════════════
var _holdingsDataMemo = null;

// ── เคลียร์ _holdingsDataMemo — เรียกทันทีหลังเขียนข้อมูลกลับลง US_HOLD/TH_HOLD
//    ในรอบ execution เดียวกัน (เช่น อัปเดตราคาแล้วต้องอ่าน holdings สดต่อทันที) ──
function _clearHoldingsDataMemo() {
  _holdingsDataMemo = null;
}


function getHoldingsData() {
  if (_holdingsDataMemo !== null) return _holdingsDataMemo;

  try {
    const usHoldings = getHoldings(SHEETS.US_HOLD);
    const thHoldings = getHoldings(SHEETS.TH_HOLD);

    const usLogRows = _getSheetDataCached(SHEETS.US_TRANS);
    const thLogRows = _getSheetDataCached(SHEETS.TH_TRANS);

    const usLowestMap = getLowestBuyPriceMap(usLogRows);
    const thLowestMap = getLowestBuyPriceMap(thLogRows);

    usHoldings.forEach(h => { h.lowestBuyPrice = usLowestMap[h.ticker] || null; });
    thHoldings.forEach(h => { h.lowestBuyPrice = thLowestMap[h.ticker] || null; });

    _holdingsDataMemo = {
      us: usHoldings,
      th: thHoldings,
      fund: getFundHoldings()
    };
    return _holdingsDataMemo;
  } catch (e) {
    logError('getHoldingsData', e);
    return { us: [], th: [], fund: [], error: e.message };
  }
}


// สร้าง map เก็บราคาต่ำสุดต่อหุ้น (ตัด _cX ทิ้งก่อน) จาก transaction log
// คอลัมน์ index (0-based) ตรงกับ saveUSStock/saveTHStock: C=หุ้น(2) D=ประเภท(3) F=ราคา(5)
function getLowestBuyPriceMap(logRows) {
  const TICKER_COL = 2;
  const TYPE_COL   = 3;
  const PRICE_COL  = 5;

  const map = {};
  logRows.forEach(row => {
    const rawTicker = row[TICKER_COL];
    const type      = row[TYPE_COL];
    const price     = parseFloat(row[PRICE_COL]);
    if (!rawTicker || type !== 'ซื้อ' || isNaN(price)) return;

    const cleanTicker = String(rawTicker).trim().toUpperCase().replace(/_C\d+$/i, '');
    if (!(cleanTicker in map) || price < map[cleanTicker]) {
      map[cleanTicker] = price;
    }
  });
  return map;
}


// ══════════════════════════════════════════════════════════
// อัปเดตราคาหุ้นล่าสุด — ปุ่ม "อัปเดตราคา" บนหน้าเว็บ
// ✅ ใช้ updateThaiPrice() / updateUSPrice() ตัวเดียวกับคำสั่ง /Updatepcs
//    ใน update.gs เดิมทั้งหมด — ต่างกันแค่ return object แทนส่ง Telegram
// ══════════════════════════════════════════════════════════
function updatePricesWeb() {
  try {
    const thCount = updateThaiPrice();
    Utilities.sleep(500);
    const usCount = updateUSPrice();
    return {
      success: true,
      th: thCount,
      us: usCount,
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (e) {
    logError('updatePricesWeb', e);
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════
// Financial Goal + CAGR — สำหรับหน้าเว็บ
// ✅ สูตรเดียวกับ sendPortfolioDashboard() ใน dashboard.gs (/dashboard เดิม)
//    ใช้ collectAllData() + getMasterData() ตัวเดียวกัน ไม่คำนวณซ้ำ
// ══════════════════════════════════════════════════════════
function getFinancialGoalData() {
  try {
    const d = collectAllData();
    const master = getMasterData();

    const totalAsset = d.totalPortTHB + d.cashTotal;
    const netInvest = d.netInvest > 0 ? d.netInvest : master.initialInvest;
    const portTarget = master.portTarget;
    const progress = portTarget > 0 ? (totalAsset / portTarget) * 100 : 0;

    const startDate = master.startDate instanceof Date ? master.startDate : new Date(master.startDate);
    const daysDiff = Math.max(1, (new Date() - startDate) / (1000 * 60 * 60 * 24));
    const years = daysDiff / 365;
    const cagr = (netInvest > 0 && years > 0)
      ? (Math.pow(totalAsset / netInvest, 1 / years) - 1) * 100
      : 0;

    return {
      success: true,
      netInvest: netInvest,
      portTarget: portTarget,
      totalAsset: totalAsset,
      progress: progress,
      years: years,
      cagr: cagr
    };
  } catch (e) {
    logError('getFinancialGoalData', e);
    return { success: false, error: e.message, netInvest: 0, portTarget: 0,
             totalAsset: 0, progress: 0, years: 0, cagr: 0 };
  }
}

// ══════════════════════════════════════════════════════════
// เขียนราคาสดกลับเข้าไปในชีต Holdings จริง (US_HOLD/TH_HOLD)
// เรียกจากหน้า Stock Detail (Portfolio Analysis) ตอนกด ▶️ auto-refresh
// เพื่อให้หน้า Holdings หลัก/Summary/Rebalance เห็นราคาใหม่ตรงกันด้วย
// ไม่ใช่แค่หน้า Stock Detail หน้าเดียว
// ══════════════════════════════════════════════════════════
function writeLivePriceToHoldingsSheet(sheetName, ticker, livePrice) {
  try {
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW.HOLD) return { success: false, error: 'ไม่มีข้อมูลในชีต' };

    const numRows = lastRow - START_ROW.HOLD + 1;
    const rows = sheet.getRange(START_ROW.HOLD, 1, numRows, 17).getValues();

    let updated = false;
    for (let i = 0; i < rows.length; i++) {
      const rowTicker = String(rows[i][HOLD_COL.TICKER - 1] || '').trim().toUpperCase();
      if (rowTicker !== ticker.toUpperCase()) continue;

      const sharesRemain = Number(rows[i][HOLD_COL.SHARES_REMAIN - 1]) || 0;
      const totalCost = Number(rows[i][HOLD_COL.TOTAL_COST - 1]) || 0;
      if (sharesRemain <= 0) continue; // ไม้ที่ขายหมดแล้ว ข้าม ไม่เขียนทับ

      const valueNow = livePrice * sharesRemain;
      const unrealizedPL = valueNow - totalCost;
      const unrealizedPct = totalCost > 0 ? (unrealizedPL / totalCost) : 0;

      const sheetRow = START_ROW.HOLD + i;
      sheet.getRange(sheetRow, HOLD_COL.PRICE_NOW).setValue(livePrice);
     // sheet.getRange(sheetRow, HOLD_COL.VALUE_NOW).setValue(valueNow);
     // sheet.getRange(sheetRow, HOLD_COL.UNREALIZED_PL).setValue(unrealizedPL);
     // sheet.getRange(sheetRow, HOLD_COL.UNREALIZED_PCT).setValue(unrealizedPct);
      updated = true;
    }

    return updated ? { success: true } : { success: false, error: 'ไม่พบแถวของ ' + ticker + ' ในชีต (อาจขายหมดแล้ว)' };
  } catch (e) {
    logError('writeLivePriceToHoldingsSheet', e);
    return { success: false, error: e.message };
  } finally {
    // ✅ เคลียร์ memo cache เสมอ ไม่ว่าจะ success หรือไม่ — กันข้อมูล PRICE_NOW เก่าค้าง
    //    ถ้ามีจุดไหนใน execution เดียวกันเรียก getHoldingsData() ต่อจากฟังก์ชันนี้
    _clearHoldingsDataMemo();
  }
}




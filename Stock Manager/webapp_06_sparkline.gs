// ============================================================
// webapp_06_sparkline.gs — กราฟเส้นเล็ก (sparkline) ในการ์ด Holdings
// อ่านราคาปิดย้อนหลังจากไฟล์ Daily_Close_Log (คนละไฟล์กับ Trading_Stock)
// ดูโครงสร้างไฟล์ทั้งหมดที่ webapp_00_main.gs
// ============================================================

const SPARKLINE_SHEET_ID = '12rlj7SR-Xofj8tdyu3atA9kUTLC44Y6Ja9cNfd2PUvw';
const SPARKLINE_TAB_NAME = 'Daily_Close_Log';
const SPARKLINE_MAX_POINTS = 20; // จำนวนจุดล่าสุดต่อหุ้นที่ใช้วาดกราฟ (ไม่ต้องเยอะ กราฟเล็ก)

// ──────────────────────────────────────
// ดึงราคาปิดย้อนหลังของทุก symbol ในชีต Daily_Close_Log
// return: { SYMBOL: [close1, close2, ..., closeล่าสุด], ... }
// เรียงจากเก่า → ใหม่ (จุดสุดท้าย = ราคาล่าสุดในชีตนี้)
// ──────────────────────────────────────
function getSparklineData() {
  try {
    const ss = SpreadsheetApp.openById(SPARKLINE_SHEET_ID);
    const sheet = ss.getSheetByName(SPARKLINE_TAB_NAME);
    if (!sheet) return {};

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    // A:Date  B:Symbol  C:Price  D:High  E:Low  F:Close
    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    const bySymbol = {};
    rows.forEach(r => {
      const date   = r[0];
      const symbol = String(r[1] || '').trim().toUpperCase();
      const close  = Number(r[5]);
      if (!symbol || !(date instanceof Date) || isNaN(close)) return;
      if (!bySymbol[symbol]) bySymbol[symbol] = [];
      bySymbol[symbol].push({ t: date.getTime(), close: close });
    });

    const result = {};
    Object.keys(bySymbol).forEach(sym => {
      const sorted = bySymbol[sym].sort((a, b) => a.t - b.t);
      const lastPoints = sorted.slice(-SPARKLINE_MAX_POINTS);
      result[sym] = lastPoints.map(p => p.close);
    });

    return result;
  } catch (e) {
    logError('getSparklineData', e);
    return {};
  }
}



// ============================================================
// webapp_24_realized_pnl_history.gs — สร้างใหม่ทั้งหมด
// เก็บ snapshot "กำไรจริง" (Realized P/L) แยกหุ้นสหรัฐ (USD) / หุ้นไทย (THB)
// ทุกวันตอนตี 4 (Asia/Bangkok) แล้วเอามาคำนวณส่วนต่างเทียบวันก่อนหน้า
// สำหรับแสดงในหน้า Summary เป็น $ / ฿ / % (+/-)
// ============================================================

const REALIZED_LOG_SHEET_NAME = 'RealizedPnL_Log';

// ── สร้าง/ดึงชีตเก็บ log ถ้ายังไม่มีให้สร้างพร้อม header ──
function _ensureRealizedPnLLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REALIZED_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(REALIZED_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['Date', 'US_Realized_USD', 'TH_Realized_THB', 'FX_Rate']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── ฟังก์ชันหลัก: บันทึก snapshot ของวันนี้ (เรียกจาก trigger ตี 4 ทุกวัน) ──
// ถ้ามี row ของ "วันนี้" (ตามเขตเวลา Asia/Bangkok) อยู่แล้ว จะอัปเดตทับแทนการเพิ่มแถวใหม่
// (กันข้อมูลซ้ำถ้ามีคนกดรันมือซ้ำ หรือ trigger ยิงซ้ำในวันเดียวกัน)
function logRealizedPnLSnapshot() {
  try {
    const realized = getRealizedPnL(); // { thNet, usNet, fundNet, totalTHB } — data.gs
    const fx = getFxRate();
    const sheet = _ensureRealizedPnLLogSheet();
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r =>
        r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[0])
      );
      const idx = dates.indexOf(today);
      if (idx !== -1) {
        sheet.getRange(idx + 2, 2, 1, 3).setValues([[realized.usNet, realized.thNet, fx]]);
        return;
      }
    }
    sheet.appendRow([new Date(), realized.usNet, realized.thNet, fx]);
  } catch (e) {
    logError('logRealizedPnLSnapshot', e);
  }
}

// ── ตั้ง trigger รายวันตอนตี 4 (Asia/Bangkok) — ลบ trigger เดิมของฟังก์ชันนี้ก่อนกันซ้ำ ──
// เรียกจากปุ่ม "เริ่มเก็บข้อมูลตอนนี้" ในหน้า UI (initRealizedPnLHistoryTracking) หรือรันเองครั้งเดียวใน Apps Script editor ก็ได้
function createRealizedPnLSnapshotTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'logRealizedPnLSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('logRealizedPnLSnapshot')
    .timeBased()
    .everyDays(1)
    .atHour(4) // ตี 4 — ตรงกับที่ขอ (หลังราคาปิดตลาดสหรัฐคืนก่อนหน้า อัปเดตครบแล้ว)
    .create();
  logRealizedPnLSnapshot(); // สร้างจุดข้อมูลแรกทันที ไม่ต้องรอถึงตี 4 พรุ่งนี้
}

// ── เรียกจากปุ่มตั้งค่าในหน้า UI ──
function initRealizedPnLHistoryTracking() {
  try {
    createRealizedPnLSnapshotTrigger();
    return { success: true };
  } catch (e) {
    logError('initRealizedPnLHistoryTracking', e);
    return { success: false, error: e.message };
  }
}

// ── ดึง log ทั้งหมด เรียงตามวันที่ ──
function getRealizedPnLHistory() {
  try {
    const sheet = _ensureRealizedPnLLogSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, points: [], note: 'ยังไม่มีข้อมูลสะสม' };

    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues()
      .filter(r => r[0])
      .map(r => ({
        date: Utilities.formatDate(new Date(r[0]), 'Asia/Bangkok', 'yyyy-MM-dd'),
        dateLabel: Utilities.formatDate(new Date(r[0]), 'Asia/Bangkok', 'dd/MM/yy'),
        usNet: Number(r[1]) || 0,
        thNet: Number(r[2]) || 0,
        fx: Number(r[3]) || 0
      }))
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    return { success: true, points: rows };
  } catch (e) {
    logError('getRealizedPnLHistory', e);
    return { success: false, error: e.message };
  }
}

// ── เปรียบเทียบวันล่าสุด vs วันก่อนหน้าที่มีข้อมูล → คืนค่าเป็น $/฿/% พร้อมเครื่องหมาย +/- ──
// เรียกจากหน้า Summary (google.script.run.getRealizedPnLChangeSummary)
function getRealizedPnLChangeSummary() {
  try {
    const history = getRealizedPnLHistory();
    if (!history.success) return { success: false, error: history.error };

    const points = history.points || [];
    if (!points.length) {
      return {
        success: true, hasEnoughData: false,
        note: 'ยังไม่มีข้อมูลสะสม — ระบบจะเริ่มเก็บอัตโนมัติทุกวันตอนตี 4 หรือกดปุ่ม "เริ่มเก็บข้อมูลตอนนี้" เพื่อสร้างจุดแรกทันที'
      };
    }

    const latest = points[points.length - 1];
    // ⚡ หา "จุดก่อนหน้าล่าสุดที่วันที่ไม่ตรงกับ latest" กันกรณี snapshot วันนี้ถูกเรียกซ้ำ
    const prev = [...points].reverse().find(p => p.date !== latest.date) || null;

    if (!prev) {
      return {
        success: true, hasEnoughData: false,
        latest,
        note: 'มีข้อมูลแค่วันเดียว — พรุ่งนี้ตี 4 จะเริ่มเห็นการเปลี่ยนแปลง'
      };
    }

    const fx = latest.fx || getFxRate(); // ใช้ fx ของวันล่าสุดแปลง USD→THB (ถ้าไม่มีค่อย fallback ไปดึงสด

    // ── หุ้นสหรัฐ (USD) ──
    const usChangeUSD = latest.usNet - prev.usNet;
    const usChangeTHB = usChangeUSD * fx;
    const usChangePct = prev.usNet !== 0 ? (usChangeUSD / Math.abs(prev.usNet)) * 100 : null;

    // ── หุ้นไทย (THB) ──
    const thChangeTHB = latest.thNet - prev.thNet;
    const thChangePct = prev.thNet !== 0 ? (thChangeTHB / Math.abs(prev.thNet)) * 100 : null;

    return {
      success: true,
      hasEnoughData: true,
      latestDate: latest.dateLabel,
      prevDate: prev.dateLabel,
      us: {
        latestUSD: latest.usNet,
        prevUSD: prev.usNet,
        changeUSD: usChangeUSD,
        changeTHB: usChangeTHB,
        changePct: usChangePct
      },
      th: {
        latestTHB: latest.thNet,
        prevTHB: prev.thNet,
        changeTHB: thChangeTHB,
        changePct: thChangePct
      }
    };
  } catch (e) {
    logError('getRealizedPnLChangeSummary', e);
    return { success: false, error: e.message };
  }
}


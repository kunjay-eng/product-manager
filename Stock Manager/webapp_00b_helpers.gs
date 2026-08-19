// ============================================================
// webapp_00b_helpers.gs — ฟังก์ชันช่วยกลางที่ไฟล์อื่นเรียกใช้ร่วมกัน
// (getSheet, logError) — ต้องมีไฟล์นี้ก่อนไฟล์ webapp_* อื่นๆ จะทำงานได้
// ============================================================

// ── ดึงชีตตามชื่อ พร้อม error ที่อ่านง่ายถ้าหาไม่เจอ ──
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('ไม่พบชีตชื่อ "' + sheetName + '" — เช็คว่าพิมพ์ชื่อชีตถูกต้อง หรือยังไม่ได้สร้างชีตนี้');
  }
  return sheet;
}

// ── log error กลาง — เขียนลง Logger เสมอ (ดูได้ที่ Execution log) ──
// พยายามเขียนลงชีต "ErrorLog" เพิ่มด้วย ถ้าไม่มีชีตนี้จะข้ามไปเงียบๆ
// (ห้ามให้ logError เองพังซ้ำ เพราะจะไปบังข้อความ error ตัวจริง)
function logError(functionName, error) {
  const msg = (error && error.message) ? error.message : String(error);
  const stack = (error && error.stack) ? error.stack : '';
  Logger.log('❌ [' + functionName + '] ' + msg + (stack ? '\n' + stack : ''));

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('ErrorLog');
    if (!logSheet) {
      logSheet = ss.insertSheet('ErrorLog');
      logSheet.appendRow(['เวลา', 'ฟังก์ชัน', 'ข้อความ error']);
      logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    logSheet.appendRow([new Date(), functionName, msg]);
  } catch (e) {
    // เงียบไว้ — ไม่ทำให้ error หลักหายไป
  }
}


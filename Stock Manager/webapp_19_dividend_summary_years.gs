// ══════════════════════════════════════════════════════════
// webapp_19_dividend_summary_years.gs
// เติมแถวปีใหม่ในตาราง "สรุปปันผลรายปี" (คอลัมน์ U:AA ในชีต SHEETS.DIV)
// ล่วงหน้าอัตโนมัติ โดยไม่กระทบแถวบันทึกปันผลดิบ (คอลัมน์ B-R) ในชีตเดียวกัน
// ------------------------------------------------------------
// หลักการ:
// 1. หาแถว "รวม" (total) ในคอลัมน์ U เพื่อกำหนดขอบเขตตาราง
// 2. ใช้ Range.insertCells(Dimension.ROWS) แทรกเฉพาะคอลัมน์ U:AA
//    (ไม่ใช้ insertRowBefore ทั้งแถว เพราะจะดันข้อมูลปันผลดิบเลื่อนไปด้วย)
// 3. คัดลอก "รูปแบบ" (format) จากแถวปีล่าสุดด้วย copyFormatToRange
//    แล้วเขียนสูตรของ V/W/X ใหม่เอง (แทนที่ตัวเลขปีในสูตรเดิม) —
//    ไม่ใช้ copyTo() เพราะจะไปเลื่อน range I7:I186/M7:M186 ที่เป็น
//    ขอบเขตตายตัวของตารางปันผลดิบ ทำให้สูตรพังได้
// 4. อัปเดตสูตร SUM ของแถว "รวม" ให้ครอบคลุมปีใหม่ที่เพิ่ม
// ══════════════════════════════════════════════════════════

const DIVSUM_COL = { YEAR: 21, THB: 22, USD: 23, TOTAL: 24, DIFF_THB: 25, DIFF_PCT: 26, TARGET: 27 }; // U..AA
const DIVSUM_YEARS_AHEAD = 5; // เผื่อล่วงหน้าเสมอ N ปีจากปีปัจจุบัน

/**
 * เติมแถวปีที่ขาดหายไปในตารางสรุปปันผลรายปี จนถึง (ปีปัจจุบัน + DIVSUM_YEARS_AHEAD)
 * เรียกได้ทั้งจากปุ่มในหน้า Settings และจาก trigger อัตโนมัติ
 */
function addMissingDividendSummaryYears() {
  try {
    const sheet = getSheet(SHEETS.DIV);
    const lastRow = sheet.getLastRow();

    // ── หาแถว "รวม" (total) และไล่เก็บแถวปีทั้งหมดที่มีอยู่แล้ว ──
    let totalRow = -1;
    const yearRows = []; // [{row, year}]
    for (let r = 1; r <= lastRow; r++) {
      const val = sheet.getRange(r, DIVSUM_COL.YEAR).getValue();
      if (String(val).trim() === 'รวม') { totalRow = r; break; }
      if (typeof val === 'number' && val >= 2000 && val <= 2100) yearRows.push({ row: r, year: val });
    }
    if (totalRow === -1 || !yearRows.length) {
      return { success: false, error: 'ไม่พบตารางสรุปปันผลรายปี (คอลัมน์ U) หรือแถว "รวม" — ตรวจโครงสร้างชีตอีกครั้ง' };
    }

    const firstYearRow = yearRows[0];
    let lastYearEntry = yearRows[yearRows.length - 1];

    const currentYear = new Date().getFullYear();
    const targetMaxYear = currentYear + DIVSUM_YEARS_AHEAD;

    const missingYears = [];
    for (let y = lastYearEntry.year + 1; y <= targetMaxYear; y++) missingYears.push(y);
    if (!missingYears.length) {
      return { success: true, added: 0, message: 'มีข้อมูลครบถึงปี ' + lastYearEntry.year + ' อยู่แล้ว (เผื่อไว้ถึงปี ' + targetMaxYear + ')' };
    }

    // ── คำนวณส่วนต่างเป้าหมายจาก 2 ปีล่าสุด ใช้ต่อยอดปีใหม่ (ถ้าคำนวณไม่ได้ ใช้ค่าเดิมซ้ำ) ──
    let targetIncrement = 0;
    if (yearRows.length >= 2) {
      const lastTarget = Number(sheet.getRange(lastYearEntry.row, DIVSUM_COL.TARGET).getValue()) || 0;
      const prevTarget = Number(sheet.getRange(yearRows[yearRows.length - 2].row, DIVSUM_COL.TARGET).getValue()) || 0;
      targetIncrement = lastTarget - prevTarget;
    }

    missingYears.forEach(year => {
      // ── แทรกเฉพาะคอลัมน์ U:AA ที่ตำแหน่ง totalRow (ดันแถว "รวม" กับปีอื่นๆ ที่อยู่ล่างมันลง 1 แถว
      //    เฉพาะในคอลัมน์นี้เท่านั้น — คอลัมน์อื่นในแถวเดียวกัน (ปันผลดิบ) ไม่ถูกแตะต้องเลย) ──
      const insertPoint = sheet.getRange(totalRow, DIVSUM_COL.YEAR, 1, 7); // U..AA = 7 คอลัมน์
      insertPoint.insertCells(SpreadsheetApp.Dimension.ROWS);

      const newRow = totalRow; // แถวว่างใหม่ที่เพิ่งแทรก อยู่ตำแหน่งเดิมของ totalRow (ก่อนถูกดันลง)

      // ── คัดลอกรูปแบบ (สี, border, number format) จากแถวปีล่าสุด — ไม่แตะสูตร/ค่า ──
      const sourceRange = sheet.getRange(lastYearEntry.row, DIVSUM_COL.YEAR, 1, 7);
      const destRange = sheet.getRange(newRow, DIVSUM_COL.YEAR, 1, 7);
      sourceRange.copyFormatToRange(sheet, DIVSUM_COL.YEAR, DIVSUM_COL.TARGET, newRow, newRow);

      // ── ปี ──
      sheet.getRange(newRow, DIVSUM_COL.YEAR).setValue(year);

      // ── สูตร ปันผลไทย (V) / ปันผลสหรัฐ (W) / รวม (X) — ดึงสูตรจากแถวปีล่าสุด
      //    แทนที่ "เลขปีเดิม" ด้วย "ปีใหม่" ในตัวสูตร แล้วเขียนกลับตรงๆ (ไม่ผ่าน copyTo
      //    เพราะ copyTo จะพยายามเลื่อน range I7:I186 ตามระยะห่างแถว ทำให้สูตรอ้างผิดช่วง) ──
      [DIVSUM_COL.THB, DIVSUM_COL.USD, DIVSUM_COL.TOTAL].forEach(col => {
        const srcFormula = sheet.getRange(lastYearEntry.row, col).getFormula();
        if (srcFormula) {
          const newFormula = srcFormula.replace(new RegExp(String(lastYearEntry.year), 'g'), String(year));
          sheet.getRange(newRow, col).setFormula(newFormula);
        }
      });

      // ── เพิ่ม/ลด (THB) และ เพิ่ม/ลด % — เทียบกับแถวปีก่อนหน้าโดยตรง (แถวติดกันเสมอ) ──
      sheet.getRange(newRow, DIVSUM_COL.DIFF_THB).setFormula(
        `=X${newRow}-X${lastYearEntry.row}`
      );
      sheet.getRange(newRow, DIVSUM_COL.DIFF_PCT).setFormula(
        `=IFERROR((X${newRow}-X${lastYearEntry.row})/X${lastYearEntry.row},"—")`
      );

      // ── เป้าหมาย — ต่อยอดจากส่วนต่าง 2 ปีล่าสุดที่คำนวณไว้ (ผู้ใช้แก้ไขเองได้ทีหลัง) ──
      const lastTargetVal = Number(sheet.getRange(lastYearEntry.row, DIVSUM_COL.TARGET).getValue()) || 0;
      sheet.getRange(newRow, DIVSUM_COL.TARGET).setValue(Math.max(0, lastTargetVal + targetIncrement));

      // ── เลื่อนตัวชี้ไปแถวใหม่ สำหรับวนรอบปีถัดไป ──
      lastYearEntry = { row: newRow, year };
      totalRow = newRow + 1; // แถว "รวม" ถูกดันลงไป 1 แถวเสมอในแต่ละรอบ
    });

    // ── อัปเดตสูตร SUM ของแถว "รวม" ให้ครอบคลุมช่วงปีใหม่ทั้งหมด (first...last) ──
    [DIVSUM_COL.THB, DIVSUM_COL.USD, DIVSUM_COL.TOTAL].forEach(col => {
      const totalFormula = sheet.getRange(totalRow, col).getFormula();
      if (totalFormula) {
        const colLetter = String.fromCharCode(64 + col); // 22->V, 23->W, 24->X
        const newFormula = totalFormula.replace(
          new RegExp(colLetter + '\\d+:' + colLetter + '\\d+', 'g'),
          colLetter + firstYearRow.row + ':' + colLetter + lastYearEntry.row
        );
        sheet.getRange(totalRow, col).setFormula(newFormula);
      }
    });

    return {
      success: true,
      added: missingYears.length,
      years: missingYears,
      message: 'เพิ่มปี ' + missingYears.join(', ') + ' แล้ว (' + missingYears.length + ' ปี)'
    };
  } catch (e) {
    logError('addMissingDividendSummaryYears', e);
    return { success: false, error: e.message };
  }
}

// ── ตั้ง trigger อัตโนมัติ: เช็คทุกวันที่ 1 ของเดือน เวลาตี 2 — ถ้าปีล่าสุดในตาราง
//    ใกล้หมด (น้อยกว่า currentYear+DIVSUM_YEARS_AHEAD) จะเติมให้เองอัตโนมัติ
//    ไม่ต้องรอผู้ใช้กดปุ่มเอง (self-healing เหมือน logPortfolioValueSnapshot) ──
function createDividendSummaryYearTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'addMissingDividendSummaryYears') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('addMissingDividendSummaryYears').timeBased().onMonthDay(1).atHour(2).create();
  addMissingDividendSummaryYears(); // รันทันที 1 ครั้งตอนตั้งค่า กันรอถึงรอบเดือนหน้า
}

// ── รันเองใน Apps Script Editor เพื่อทดสอบ ──
function testAddMissingDividendSummaryYears() {
  Logger.log(JSON.stringify(addMissingDividendSummaryYears(), null, 2));
}


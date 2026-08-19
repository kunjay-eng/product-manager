// ============================================================
// webapp_09_entry_extra.gs — บันทึกธุรกรรมส่วนที่เหลือ (นอกจาก US/TH stock, ปันผล
// ที่มีอยู่แล้วใน Code.gs): กองทุน, ฝาก/ถอน, แลกเงิน, อัปเดต NAV, ตั้งค่า FX Alert
//
// ⚠️ หมายเหตุสำคัญ: ชื่อชีตด้านล่าง (FUND_TRANS_SHEET, FX_EXCHANGE_SHEET,
// FX_ALERT_SHEET) เป็นชื่อที่ประมาณจากภาพหน้าจอที่ส่งมา — ถ้าไม่ตรงกับชื่อ
// แท็บจริงในไฟล์ Trading_Stock ให้แก้ค่า 3 บรรทัดนี้ให้ตรงเท่านั้น
// (ส่วนอื่นของไฟล์ไม่ต้องแก้)
// ============================================================


const FUND_TRANS_SHEET  = '🏛️ กองทุนรวม';   // ⚠️ เช็คชื่อแท็บจริง
const FX_EXCHANGE_SHEET = 'ข้อมูลการแลกเงิน';   // ยืนยันจากภาพที่ส่งมาแล้ว
const FX_ALERT_SHEET    = 'FX_Alert';           // ⚠️ เช็คชื่อแท็บจริง
const CASH_LOG_START_ROW = 18; // แถวแรกของตารางรายการฝาก/ถอน (ดูจากภาพ Cash Flow)
const FUND_TRANS_START_ROW = 5;
const FX_EXCHANGE_START_ROW = 2;


// ══════════════════════════════════════════════════════════
// 🏦 ซื้อ/ขายกองทุน — col B,C,D,E,F,H,I,J,L (K=ยอดสุทธิ คำนวณเป็นสูตร)
// B=วันที่ C=ชื่อกองทุน D=ประเภท E=จำนวนหน่วย F=NAV/หน่วย
// H=ค่าธรรมเนียมซื้อ I=ค่าธรรมเนียมขาย J=ภาษีหัก ณ ที่จ่าย L=NAV ปัจจุบัน(ล่าสุด, ถ้ามี)
// ══════════════════════════════════════════════════════════
function saveFundTransaction(data) {
  try {
    const sheet = getSheet(FUND_TRANS_SHEET);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + FUND_TRANS_SHEET);


    const row = _getNextEmptyRow(sheet, 3, FUND_TRANS_START_ROW); // เช็คจาก C (ชื่อกองทุน)
    const dateVal = new Date(data.date);


    sheet.getRange(row, 2).setValue(dateVal).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 3).setValue(data.fundName);
    sheet.getRange(row, 4).setValue(data.type); // ซื้อ / ขาย
    sheet.getRange(row, 5).setValue(parseFloat(data.units) || 0);
    sheet.getRange(row, 6).setValue(parseFloat(data.nav) || 0);
    sheet.getRange(row, 7).setFormula(`=E${row}*F${row}`); // มูลค่ารวม
    sheet.getRange(row, 8).setValue(parseFloat(data.feeBuy) || 0);
    sheet.getRange(row, 9).setValue(parseFloat(data.feeSell) || 0);
    sheet.getRange(row, 10).setValue(parseFloat(data.tax) || 0);
    // ยอดสุทธิ: ซื้อ = มูลค่า+ค่าธรรมเนียมซื้อ+ภาษี | ขาย = มูลค่า−ค่าธรรมเนียมขาย−ภาษี
    sheet.getRange(row, 11).setFormula(
      `=IF(D${row}="ซื้อ", G${row}+H${row}+J${row}, G${row}-I${row}-J${row})`
    );
    if (data.currentNav) {
      sheet.getRange(row, 12).setValue(parseFloat(data.currentNav) || 0); // L (ถ้ากรอกมา)
    }


    sheet.getRange(row, 5).setNumberFormat('0.000000');
    sheet.getRange(row, 6, 1, 6).setNumberFormat('#,##0.00');


    return { success: true, row: row };
  } catch (e) {
    logError('saveFundTransaction', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// 🏦 ฝาก/ถอนเงิน — col B(#auto),C(วันที่),D(จำนวน THB),E(หมายเหตุ),F(ประเภท)
// เขียนต่อท้ายตาราง Cash Flow (แถวเริ่ม CASH_LOG_START_ROW)
// ══════════════════════════════════════════════════════════
function saveCashFlow(data) {
  try {
    const sheet = getSheet(SHEETS.CASH);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SHEETS.CASH);


    const row = _getNextEmptyRow(sheet, 3, CASH_LOG_START_ROW); // เช็คจาก C (วันที่)
    const lastRow = sheet.getLastRow();
    const seq = row - CASH_LOG_START_ROW + 1; // ลำดับที่ #


    sheet.getRange(row, 2).setValue(seq);
    sheet.getRange(row, 3).setValue(new Date(data.date)).setNumberFormat('dd-mm-yyyy');
    sheet.getRange(row, 4).setValue(parseFloat(data.amount) || 0).setNumberFormat('#,##0.00');
    sheet.getRange(row, 5).setValue(data.note || '');
    sheet.getRange(row, 6).setValue(data.type); // Deposit / Withdrawal


    return { success: true, row: row };
  } catch (e) {
    logError('saveCashFlow', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// 💱 แลกเปลี่ยนเงิน — col A(ลำดับ auto),B(วันที่),C(เวลา),D(ประเภท),
// E(จำนวนเงินต้นทาง กรอกเอง), G(จำนวนเงินปลายทาง กรอกเอง),
// I(อัตราแลกเปลี่ยน กรอกเอง)
// F/H (ป้ายชื่อสกุลเงินต้นทาง/ปลายทาง) คำนวณเป็นสูตรจาก D ให้อัตโนมัติ
// หมายเหตุ: E และ G กรอกเป็นตัวเลขจริงที่ได้รับทั้งคู่ (ไม่คำนวณจากกันเอง)
// เพราะยอดจริงอาจเพี้ยนจากค่าตามทฤษฎีเล็กน้อยจากค่าธรรมเนียม/การปัดเศษ
// ══════════════════════════════════════════════════════════
function saveCurrencyExchange(data) {
  try {
    const sheet = getSheet(FX_EXCHANGE_SHEET);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + FX_EXCHANGE_SHEET);


    const row = _getNextEmptyRow(sheet, 2, FX_EXCHANGE_START_ROW); // เช็คจาก B (วันที่)
    const seq = row - FX_EXCHANGE_START_ROW + 1;


    sheet.getRange(row, 1).setValue(seq);
    sheet.getRange(row, 2).setValue(new Date(data.date)).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(row, 3).setValue(data.time || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss'));
    sheet.getRange(row, 4).setValue(data.direction); // "THB→USD" หรือ "USD→THB"
    sheet.getRange(row, 5).setValue(parseFloat(data.sourceAmount) || 0); // E: จำนวนเงินต้นทาง (กรอกเอง)
    sheet.getRange(row, 7).setValue(parseFloat(data.destAmount) || 0);   // G: จำนวนเงินปลายทาง (กรอกเอง)
    sheet.getRange(row, 9).setValue(parseFloat(data.rate) || 0);        // I: อัตราแลกเปลี่ยน (กรอกเอง)


    // F/H: ป้ายชื่อสกุลเงิน (ต้นทาง/ปลายทาง) ตามทิศทาง — เผื่อแถวนี้ยังไม่มีสูตรเดิม
    sheet.getRange(row, 6).setFormula(`=IF(D${row}="THB→USD","THB",IF(D${row}="USD→THB","USD",""))`);
    sheet.getRange(row, 8).setFormula(`=IF(D${row}="THB→USD","USD",IF(D${row}="USD→THB","THB",""))`);


    sheet.getRange(row, 5).setNumberFormat('#,##0.00');
    sheet.getRange(row, 7).setNumberFormat('#,##0.00');
    sheet.getRange(row, 9).setNumberFormat('0.00');


    return { success: true, row: row };
  } catch (e) {
    logError('saveCurrencyExchange', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// 📊 บันทึก NAV ล่าสุด — อัปเดตคอลัมน์ H ในชีต Holdings กองทุน
// (SHEETS.FUND_HOLD เดิม ที่ getFundHoldings() ใช้อยู่แล้ว)
// ══════════════════════════════════════════════════════════
function updateFundNAV(data) {
  try {
    const sheet = getSheet(SHEETS.FUND_HOLD);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + SHEETS.FUND_HOLD);


    const lastRow = sheet.getLastRow();
    const startRow = START_ROW.FUND;
    if (lastRow < startRow) return { success: false, error: 'ไม่มีรายการกองทุนในชีต' };


    const names = sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues(); // col B ชื่อกองทุน
    let targetRow = -1;
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === String(data.fundName).trim()) {
        targetRow = startRow + i;
        break;
      }
    }
    if (targetRow === -1) return { success: false, error: 'ไม่พบกองทุน "' + data.fundName + '" ในชีต Holdings' };


    sheet.getRange(targetRow, 8).setValue(parseFloat(data.nav) || 0).setNumberFormat('0.0000'); // H
    return { success: true, row: targetRow };
  } catch (e) {
    logError('updateFundNAV', e);
    return { success: false, error: e.message };
  }

}


// ══════════════════════════════════════════════════════════
// 🔔 ตั้งค่า FX Alert — อ่าน/เขียน B2:B4 (Alert_Buy_1, Alert_Buy_2, Alert_Sell)
// B1 (USDTHB) เป็นค่าที่ระบบอื่นอัปเดตอัตโนมัติอยู่แล้ว ไม่เขียนทับจากหน้านี้
// ══════════════════════════════════════════════════════════
function getFXAlertSettings() {
  try {
    const sheet = getSheet(FX_ALERT_SHEET);
    return {
      success: true,
      usdthb: Number(sheet.getRange('B1').getValue()) || 0,
      alertBuy1: Number(sheet.getRange('B2').getValue()) || 0,
      alertBuy2: Number(sheet.getRange('B3').getValue()) || 0,
      alertSell: Number(sheet.getRange('B4').getValue()) || 0
    };
  } catch (e) {
    logError('getFXAlertSettings', e);
    return { success: false, error: e.message };
  }
}


function saveFXAlertSettings(data) {
  try {
    const sheet = getSheet(FX_ALERT_SHEET);
    sheet.getRange('B2').setValue(parseFloat(data.alertBuy1) || 0);
    sheet.getRange('B3').setValue(parseFloat(data.alertBuy2) || 0);
    sheet.getRange('B4').setValue(parseFloat(data.alertSell) || 0);
    return { success: true };
  } catch (e) {
    logError('saveFXAlertSettings', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// ✅ Validation ก่อนขาย — เช็คว่ามีหุ้น/หน่วยกองทุนพอขายหรือไม่
// เรียกจากหน้าเว็บตอนกรอกฟอร์ม "ขาย" เพื่อโชว์จำนวนคงเหลือ + ปุ่ม "ขายทั้งหมด"
// ใช้ getHoldings()/getFundHoldings() เดิม ไม่คำนวณ position ซ้ำ
// ══════════════════════════════════════════════════════════
function getSellableQty(ticker, assetType) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    let rows = [];
    if (assetType === 'fund') {
      rows = getFundHoldings();
      const match = rows.find(r => String(r.name).trim().toUpperCase() === ticker);
      return { success: true, qty: match ? match.unitsRemain : 0, found: !!match };
    }
    const market = assetType === 'th' ? SHEETS.TH_HOLD : SHEETS.US_HOLD;
    rows = getHoldings(market);
    const match = rows.find(r => String(r.ticker).trim().toUpperCase() === ticker);
    return { success: true, qty: match ? match.sharesRemain : 0, found: !!match };
  } catch (e) {
    logError('getSellableQty', e);
    return { success: false, error: e.message, qty: 0, found: false };
  }
}






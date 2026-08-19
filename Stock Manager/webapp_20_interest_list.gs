// ══════════════════════════════════════════════════════════
// webapp_20_interest_list.gs — 💡 หุ้นที่สนใจ (Interest List)
// ------------------------------------------------------------
// รายชื่อหุ้นที่อยากเก็บไว้ดูก่อน ยังไม่ต้องวิเคราะห์อะไรเลย
// เก็บแค่ Ticker + ชื่อบริษัท (ดึงจาก Yahoo ครั้งเดียวตอนเพิ่ม ไม่คำนวณ
// technical ใดๆ) แยกหุ้นไทย/สหรัฐ ลบได้ตรงๆ (hard delete ไม่ต้อง soft
// delete เพราะเป็นแค่ลิสต์เก็บไอเดีย ไม่มีประวัติที่ต้องรักษาไว้)
//
// ⚠️ reuse _wlFetchYahooQuote() จาก webapp_07_watchlist.gs (โปรเจกต์
//    เดียวกัน ไม่ต้องประกาศซ้ำ) เพื่อดึงชื่อบริษัทตอนเพิ่ม
// ══════════════════════════════════════════════════════════

const INTEREST_SHEET_NAME = '💡 หุ้นที่สนใจ';
const INTEREST_START_ROW = 2;
// คอลัมน์: A=Ticker B=Market C=CompanyName D=Note E=DateAdded

function _getOrCreateInterestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INTEREST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INTEREST_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['Ticker', 'Market', 'CompanyName', 'Note', 'DateAdded']])
      .setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 5, 140);
  }
  return sheet;
}

// ── ดึงรายการทั้งหมด แยกกลุ่มตามตลาดให้ frontend ──
function getInterestListData() {
  try {
    const sheet = _getOrCreateInterestSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < INTEREST_START_ROW) return { success: true, us: [], th: [] };

    const numRows = lastRow - INTEREST_START_ROW + 1;
    const rows = sheet.getRange(INTEREST_START_ROW, 1, numRows, 5).getValues();

    const us = [], th = [];
    rows.forEach((r, i) => {
      const ticker = String(r[0] || '').trim().toUpperCase();
      if (!ticker) return;
      const market = String(r[1] || 'US').trim().toUpperCase();
      const item = {
        rowIndex: INTEREST_START_ROW + i,
        ticker,
        market,
        companyName: r[2] || '',
        note: r[3] || '',
        dateAdded: r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Bangkok', 'dd/MM/yyyy') : ''
      };
      (market === 'TH' ? th : us).push(item);
    });

    return { success: true, us, th };
  } catch (e) {
    logError('getInterestListData', e);
    return { success: false, error: e.message, us: [], th: [] };
  }
}

// ── เพิ่มหุ้นเข้าลิสต์ — ดึงชื่อบริษัทอัตโนมัติจาก Yahoo (ครั้งเดียว ไม่คำนวณ technical ใดๆ) ──
function addInterestListItem(ticker, market, note) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = (String(market || '').trim().toUpperCase() === 'TH') ? 'TH' : 'US';
    if (!ticker) return { success: false, error: 'กรุณาระบุ ticker' };

    // ── กันเพิ่มซ้ำ ──
    const existing = getInterestListData();
    const dup = [...(existing.us || []), ...(existing.th || [])].some(x => x.ticker === ticker && x.market === market);
    if (dup) return { success: false, error: ticker + ' อยู่ในลิสต์นี้อยู่แล้ว' };

    const quote = _wlFetchYahooQuote(ticker, market); // reuse จาก webapp_07_watchlist.gs
    if (!quote) return { success: false, error: 'ไม่พบหุ้น "' + ticker + '" ตรวจสอบชื่อ ticker และตลาดอีกครั้ง' };

    const sheet = _getOrCreateInterestSheet();
    const row = sheet.getLastRow() + 1;

    sheet.getRange(row, 1).setValue(ticker);
    sheet.getRange(row, 2).setValue(market);
    sheet.getRange(row, 3).setValue(quote.longName || ticker);
    sheet.getRange(row, 4).setValue(note || '');
    sheet.getRange(row, 5).setValue(new Date()).setNumberFormat('yyyy-mm-dd');

    return { success: true, row, companyName: quote.longName || ticker };
  } catch (e) {
    logError('addInterestListItem', e);
    return { success: false, error: e.message };
  }
}

// ── ลบออกจากลิสต์ (hard delete) ──
function removeInterestListItem(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = String(market || '').trim().toUpperCase();

    const sheet = _getOrCreateInterestSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < INTEREST_START_ROW) return { success: false, error: 'ไม่พบรายการ' };

    const numRows = lastRow - INTEREST_START_ROW + 1;
    const rows = sheet.getRange(INTEREST_START_ROW, 1, numRows, 2).getValues();

    for (let i = 0; i < rows.length; i++) {
      const rTicker = String(rows[i][0] || '').trim().toUpperCase();
      const rMarket = String(rows[i][1] || '').trim().toUpperCase();
      if (rTicker === ticker && rMarket === market) {
        sheet.deleteRow(INTEREST_START_ROW + i);
        return { success: true, ticker };
      }
    }
    return { success: false, error: 'ไม่พบ ' + ticker + ' ในลิสต์' };
  } catch (e) {
    logError('removeInterestListItem', e);
    return { success: false, error: e.message };
  }
}

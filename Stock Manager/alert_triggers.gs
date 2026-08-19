// ============================================================
// alert_triggers.gs — แจ้งเตือนอัตโนมัติผ่าน Telegram
// 1) Watchlist เข้า Buy Zone
// 2) หุ้นที่ถืออยู่ (Fast mode) หลุด Hard Stop
// ไม่อิงฟังก์ชัน Telegram เดิม — เขียนส่งข้อความเองใหม่ทั้งหมด
// ⚠️ ต้องรัน setupProperties() (ของเดิมที่มี BOT_TOKEN/CHAT_ID) ไว้แล้ว
// ============================================================

const ALERT_DEDUPE_SHEET_NAME = 'Alert_Dedupe_Log'; // สร้างอัตโนมัติถ้ายังไม่มี

// ══════════════════════════════════════════════════════════
// ส่งข้อความ Telegram — ฟังก์ชันใหม่ ไม่อิงของเดิม
// ══════════════════════════════════════════════════════════
function _sendAlertTelegramMessage(text) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('BOT_TOKEN');
    const chatId = props.getProperty('CHAT_ID');
    if (!token || !chatId) {
      logError('_sendAlertTelegramMessage', new Error('ไม่พบ BOT_TOKEN หรือ CHAT_ID ใน Script Properties — รัน setupProperties() ก่อน'));
      return false;
    }

    const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML'
    };
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) {
      logError('_sendAlertTelegramMessage', new Error('HTTP ' + code + ': ' + resp.getContentText()));
      return false;
    }
    return true;
  } catch (e) {
    logError('_sendAlertTelegramMessage', e);
    return false;
  }
}

// ══════════════════════════════════════════════════════════
// กันเตือนซ้ำ — เตือนได้แค่ 1 ครั้ง/วัน/เงื่อนไข (key = ticker+market+type+date)
// จนกว่าจะออกจากเงื่อนไขนั้นแล้วกลับเข้ามาใหม่ในวันถัดไปถึงจะเตือนอีกได้
// ══════════════════════════════════════════════════════════
function _ensureAlertDedupeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ALERT_DEDUPE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ALERT_DEDUPE_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['Key', 'Ticker', 'Type', 'DateSent']]);
  }
  return sheet;
}

// หมายเหตุ: ห้ามเทียบ r[3] === today ตรงๆ เพราะ Google Sheets จะ auto-convert
// สตริงรูปแบบวันที่ (เช่น "2026-07-25") ที่เขียนด้วย appendRow ให้กลายเป็น
// Date object โดยอัตโนมัติ ทำให้เทียบกับ string ไม่ตรงกันเสมอ (บั๊กเดิม
// ที่ทำให้ dedupe ไม่ทำงาน — เตือนซ้ำได้ทุกครั้งที่ trigger รัน)
function _dedupeCellToDateStr(cellVal) {
  if (cellVal instanceof Date) {
    return Utilities.formatDate(cellVal, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return String(cellVal || '');
}

function _hasAlertBeenSentToday(key) {
  const sheet = _ensureAlertDedupeSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return rows.some(r => r[0] === key && _dedupeCellToDateStr(r[3]) === today);
}

function _markAlertSent(key, ticker, type) {
  const sheet = _ensureAlertDedupeSheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const row = sheet.getLastRow() + 1;
  // ล็อกคอลัมน์ D เป็น Plain Text ก่อนเขียน กัน Sheets แปลงเป็น Date อัตโนมัติอีก
  sheet.getRange(row, 4).setNumberFormat('@');
  sheet.getRange(row, 1, 1, 4).setValues([[key, ticker, type, today]]);
}

// ── ล้าง log เก่ากว่า 7 วัน กันชีตบวม เรียกจากในตัวเช็คทุกครั้ง ──
function _cleanupOldAlertDedupeRows() {
  const sheet = _ensureAlertDedupeSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const kept = rows.filter(r => {
    const d = new Date(r[3]);
    return !isNaN(d) && d >= cutoff;
  });

  if (kept.length !== rows.length) {
    sheet.getRange(2, 1, rows.length, 4).clearContent();
    if (kept.length > 0) sheet.getRange(2, 1, kept.length, 4).setValues(kept);
  }
}

// ══════════════════════════════════════════════════════════
// ฟังก์ชัน 1: เช็ค Watchlist ที่เข้า Buy Zone แล้วยิงแจ้งเตือน
// เงื่อนไข: currentPrice <= targetPrice (zone === 'ready')
// ══════════════════════════════════════════════════════════
function checkWatchlistBuyZoneAlerts() {
  try {
    _cleanupOldAlertDedupeRows();

    const data = getWatchlistData();
    if (!data || !data.success) {
      logError('checkWatchlistBuyZoneAlerts', new Error('getWatchlistData ล้มเหลว'));
      return;
    }

    let sentCount = 0;
    (data.items || []).forEach(item => {
      if (item.zone !== 'ready') return; // ไม่เข้า Buy Zone ข้ามไป

      const key = 'BUYZONE_' + item.ticker + '_' + item.market;
      if (_hasAlertBeenSentToday(key)) return; // เตือนไปแล้ววันนี้ ข้าม

      const flag = item.market === 'TH' ? '🇹🇭' : '🇺🇸';
      const cur = item.market === 'TH' ? '฿' : '$';
      const msg =
        `🟢 <b>${item.ticker}</b> ${flag} เข้าโซนซื้อแล้ว!\n\n` +
        `ราคาปัจจุบัน: ${cur}${item.currentPrice.toFixed(2)}\n` +
        `ราคาเป้าหมาย: ${cur}${item.targetPrice.toFixed(2)}\n` +
        (item.supportPrice ? `แนวรับ: ${cur}${item.supportPrice.toFixed(2)}\n` : '') +
        `\n📱 เปิดแอปเพื่อดูรายละเอียด/วิเคราะห์เพิ่มเติม`;

      if (_sendAlertTelegramMessage(msg)) {
        _markAlertSent(key, item.ticker, 'BUYZONE');
        sentCount++;
      }
      Utilities.sleep(300); // กันยิง Telegram ถี่เกินไป
    });

    Logger.log('✅ checkWatchlistBuyZoneAlerts เสร็จสิ้น — ส่งแจ้งเตือน ' + sentCount + ' รายการ');
  } catch (e) {
    logError('checkWatchlistBuyZoneAlerts', e);
  }
}

// ══════════════════════════════════════════════════════════
// ฟังก์ชัน 2: เช็คหุ้นที่ถืออยู่ (โหมด Fast) ว่าหลุด Hard Stop หรือยัง
// ใช้ getStockDetailData() ตัวเดิม (lane='fast') เพื่อดึง tradePlan.hardStop
// ที่คำนวณตาม Risk Management ไว้แล้ว — ไม่คำนวณ Hard Stop เองใหม่
// ══════════════════════════════════════════════════════════
function checkHoldingsStopLossAlerts() {
  try {
    _cleanupOldAlertDedupeRows();

    const holdings = getHoldingsData();
    if (!holdings) {
      logError('checkHoldingsStopLossAlerts', new Error('getHoldingsData ล้มเหลว'));
      return;
    }

    const settings = getSettingsData();
    const stockModes = (settings && settings.success) ? settings.stockModes : {};

    let sentCount = 0;
    const allHoldings = [
      ...(holdings.us || []).map(r => ({ ...r, market: 'US' })),
      ...(holdings.th || []).map(r => ({ ...r, market: 'TH' }))
    ];

    allHoldings.forEach(row => {
      const ticker = row.ticker;
      if (!ticker) return;

      const qty = row.sharesRemain;
      if (!qty || qty <= 0) return; // ขายหมดแล้ว ข้าม

      const cfg = stockModes[ticker];
      if (!cfg || cfg.mode !== 'Fast') return; // เช็คเฉพาะโหมด Fast (มี Hard Stop ชัดเจน)

      const detail = getStockDetailData(ticker, row.market);
      if (!detail || !detail.success || detail.lane !== 'fast') return;

      const hardStop = detail.tradePlan && detail.tradePlan.hardStop;
      const currentPrice = detail.riskManagement && detail.riskManagement.currentPrice;
      if (!hardStop || !currentPrice) return;

      if (currentPrice >= hardStop) return; // ยังไม่หลุด Stop

      const key = 'STOPLOSS_' + ticker + '_' + row.market;
      if (_hasAlertBeenSentToday(key)) return; // เตือนไปแล้ววันนี้ ข้าม

      const flag = row.market === 'TH' ? '🇹🇭' : '🇺🇸';
      const cur = row.market === 'TH' ? '฿' : '$';
      const pl = parseFloat(row.unrealizedPL) || 0;
      const pct = (parseFloat(row.unrealizedPct) || 0) * 100;

      const msg =
        `🔴 <b>${ticker}</b> ${flag} หลุด Hard Stop แล้ว!\n\n` +
        `ราคาปัจจุบัน: ${cur}${currentPrice.toFixed(2)}\n` +
        `Hard Stop: ${cur}${hardStop.toFixed(2)}\n` +
        `กำไร/ขาดทุน: ${pl >= 0 ? '+' : ''}${fmtNumForAlert(pl)} (${pct >= 0 ? '+' : ''}${pct.toFixed(2)}%)\n` +
        `\n⚠️ ระบบ Fast Trade แนะนำให้ยึดวินัย Stop ก่อน — เปิดแอปเพื่อดูรายละเอียด/คำนวณการขาย`;

      if (_sendAlertTelegramMessage(msg)) {
        _markAlertSent(key, ticker, 'STOPLOSS');
        sentCount++;
      }
      Utilities.sleep(300);
    });

    Logger.log('✅ checkHoldingsStopLossAlerts เสร็จสิ้น — ส่งแจ้งเตือน ' + sentCount + ' รายการ');
  } catch (e) {
    logError('checkHoldingsStopLossAlerts', e);
  }
}

// ── format ตัวเลขแบบง่ายๆ สำหรับข้อความ Telegram (กันพึ่ง fmtNum ฝั่ง client) ──
function fmtNumForAlert(n) {
  return Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ══════════════════════════════════════════════════════════
// ตั้ง Trigger — เช็คทุก 1 ชั่วโมง (ปรับความถี่ได้ที่ .everyHours())
// ══════════════════════════════════════════════════════════
function createAlertCheckTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'checkWatchlistBuyZoneAlerts' || fn === 'checkHoldingsStopLossAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('checkWatchlistBuyZoneAlerts').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('checkHoldingsStopLossAlerts').timeBased().everyHours(1).create();

  Logger.log('✅ ตั้ง trigger แจ้งเตือนแล้ว — เช็คทุก 1 ชั่วโมง');
}

// ── ทดสอบยิงข้อความก่อนตั้ง trigger จริง ──
function testSendAlertMessage() {
  const ok = _sendAlertTelegramMessage('🧪 ทดสอบระบบแจ้งเตือน — ถ้าเห็นข้อความนี้แปลว่าตั้งค่าถูกต้อง');
  Logger.log(ok ? '✅ ส่งสำเร็จ' : '❌ ส่งไม่สำเร็จ ดู log ด้านบน');
}

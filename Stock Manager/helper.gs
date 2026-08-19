// ========================================
// helper.gs
// ========================================

// ----------------------------------------
// Null Check
// ----------------------------------------
function isEmpty(n) {
  return n === "" || n === null || n === undefined ||
    (typeof n === "number" && isNaN(n));
}

// ----------------------------------------
// Format Number
// ----------------------------------------
function fmt(n) {
  if (isEmpty(n)) return "-";
  return Number(n).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function fmt0(n) {
  if (isEmpty(n)) return "-";
  return Number(n).toLocaleString("en-US", {
    maximumFractionDigits: 0
  });
}

function fmt4(n) {
  if (isEmpty(n)) return "-";
  return Number(n).toLocaleString("en-US", {
    minimumFractionDigits: 4,
    maximumFractionDigits: 4
  });
}

// จำนวนหุ้น — รองรับ fractional shares
function fmtShares(n) {
  if (isEmpty(n)) return "0";
  const num = Number(n);
  if (num % 1 === 0) return fmt0(num);
  return num.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 4
  });
}

// ----------------------------------------
// Format Currency
// ----------------------------------------
function fmtTHB(n) {
  if (isEmpty(n)) return "-";
  return "฿" + fmt(n);
}

function fmtUSD(n) {
  if (isEmpty(n)) return "-";
  return "$" + fmt(n);
}

function fmtPct(n) {
  if (isEmpty(n)) return "-";
  return fmt(n) + "%";
}

// ----------------------------------------
// Sign (+/-)
// ----------------------------------------
function sign(n) {
  if (isEmpty(n)) return "-";
  return Number(n) >= 0 ? "+" + fmt(n) : fmt(n);
}

function signTHB(n) {
  if (isEmpty(n)) return "-";
  const abs = Math.abs(Number(n));
  return Number(n) >= 0 ? "+฿" + fmt(abs) : "-฿" + fmt(abs);
}

function signUSD(n) {
  if (isEmpty(n)) return "-";
  const abs = Math.abs(Number(n));
  return Number(n) >= 0 ? "+$" + fmt(abs) : "-$" + fmt(abs);
}

function signPct(n) {
  if (isEmpty(n)) return "-";
  return Number(n) >= 0 ? "+" + fmt(n) + "%" : fmt(n) + "%";
}

// ----------------------------------------
// P&L Emoji
// ----------------------------------------
function plEmoji(n) {
  if (isEmpty(n) || Number(n) === 0) return "➖";
  return Number(n) > 0 ? "🟢" : "🔴";
}

// ----------------------------------------
// Progress Bar
// ----------------------------------------
function buildProgressBar(pct) {
  const filled = Math.min(Math.round(pct / 10), 10);
  return "▓".repeat(filled) + "░".repeat(10 - filled);
}

// ----------------------------------------
// Date & Time
// ----------------------------------------
function getNow() {
  return Utilities.formatDate(
    new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss"
  );
}

function getToday() {
  return Utilities.formatDate(
    new Date(), "Asia/Bangkok", "dd/MM/yyyy"
  );
}

function formatDate(date, pattern) {
  if (!date) return "-";
  return Utilities.formatDate(
    new Date(date), "Asia/Bangkok", pattern || "dd/MM/yyyy"
  );
}

// ========================================
// helper.gs — Telegram helpers
// ========================================


// ── Token helper (single source of truth) ─
function _getTelegramCreds() {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty("BOT_TOKEN");
  const chatId = props.getProperty("CHAT_ID");
  if (!token || !chatId) {
    throw new Error("BOT_TOKEN หรือ CHAT_ID ไม่ได้ตั้งค่า — รัน setupProperties() ก่อน");
  }
  return { token, chatId };
}

// ========================================
// helper.gs — Telegram
// ========================================
function sendTelegramSafe(msg, type, symbol) {
  try {
    const { token, chatId } = _getTelegramCreds();
    if (!token || !chatId) {
      Logger.log("[ERROR] BOT_TOKEN หรือ CHAT_ID ไม่ได้ตั้งค่า");
      return;
    }
    const url = "https://api.telegram.org/bot" + token + "/sendMessage";
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: msg }),
      muteHttpExceptions: true
    });
    if (type) _logAlert(type, symbol || "", msg);
  } catch (e) {
    Logger.log("[sendTelegramSafe ERROR] " + e.message);
  }
}

function sendTelegramError(fnName, error) {
  logError(fnName, error);
  try {
    const { token, chatId } = _getTelegramCreds();
    const url = "https://api.telegram.org/bot" + token + "/sendMessage";
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: "[ERROR] " + fnName + " : " + error.message
      })
    });
  } catch (e) {
    Logger.log("[sendTelegramError ERROR] " + e.message);
  }
}








// ----------------------------------------
// Sheet Helpers
// ----------------------------------------



function getCell(sheetName, cellA1) {
  return getSheet(sheetName).getRange(cellA1).getValue();
}

function setCell(sheetName, cellA1, value) {
  getSheet(sheetName).getRange(cellA1).setValue(value);
}


// ========================================
// _getSheetDataCached() — cache ระดับ 1 execution สำหรับ getDataRange().getValues()
// ========================================
var _sheetDataCache = {};

function _getSheetDataCached(sheetName) {
  if (!(sheetName in _sheetDataCache)) {
    _sheetDataCache[sheetName] = getSheet(sheetName).getDataRange().getValues();
  }
  return _sheetDataCache[sheetName];
}

// ── เคลียร์ _sheetDataCache — เรียกทันทีหลังเขียนข้อมูลกลับลงชีตที่อาจถูก cache ไว้แล้ว
//    ในรอบ execution เดียวกัน (เช่น บันทึกธุรกรรมแล้วต้องอ่าน log สดต่อทันที)
//    ไม่ระบุ sheetName = เคลียร์ทั้งหมด ──
function _clearSheetDataCache(sheetName) {
  if (sheetName) {
    delete _sheetDataCache[sheetName];
  } else {
    _sheetDataCache = {};
  }
}


// ----------------------------------------
// Utility
// ----------------------------------------
function sleep(ms) {
  Utilities.sleep(ms);
}

// ========================================
// Yahoo Finance fetch พร้อม retry
// ใช้แทน UrlFetchApp.fetch ทุกที่
// ========================================
function fetchWithRetry(url, options, maxRetry) {
  maxRetry = maxRetry || 3;
  const delays = [1000, 2000, 4000]; // backoff

  for (let i = 0; i < maxRetry; i++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();

      if (code === 200) return resp;

      // 429 = rate limit — รอนานขึ้น
      if (code === 429) {
        logError("fetchWithRetry", "Rate limit (429) retry " + (i+1));
        Utilities.sleep(delays[i] * 2);
        continue;
      }

      // 5xx = server error — retry
      if (code >= 500) {
        logError("fetchWithRetry", "Server error (" + code + ") retry " + (i+1));
        Utilities.sleep(delays[i]);
        continue;
      }

      // 4xx อื่นๆ — ไม่ retry
      logError("fetchWithRetry", "HTTP " + code + " — no retry");
      return null;

    } catch (e) {
      logError("fetchWithRetry attempt " + (i+1), e);
      if (i < maxRetry - 1) Utilities.sleep(delays[i]);
    }
  }

  logError("fetchWithRetry", "Failed after " + maxRetry + " attempts: " + url);
  return null;
}

// ========================================
// fetch Yahoo Finance ราคาปัจจุบัน
// ========================================
function fetchYahooPrice(symbol) {
  const url  = "https://query1.finance.yahoo.com/v8/finance/chart/" +
    symbol + "?interval=1d&range=1d";
  const opts = {
    muteHttpExceptions: true,
    headers: { "User-Agent": "Mozilla/5.0" }
  };

  const resp = fetchWithRetry(url, opts, 3);
  if (!resp) return null;

  try {
    const data   = JSON.parse(resp.getContentText());
    const result = data?.chart?.result?.[0];
    if (!result) return null;

    const quote = result.indicators?.quote?.[0];
    return {
      price:  result.meta?.regularMarketPrice        || 0,
      high:   quote?.high?.[quote.high.length - 1]   || result.meta?.regularMarketDayHigh || 0,
      low:    quote?.low?.[quote.low.length - 1]      || 0,
      open:   quote?.open?.[quote.open.length - 1]    || 0,
      close:  quote?.close?.[quote.close.length - 1]  || 0,
      volume: quote?.volume?.[quote.volume.length - 1]|| 0
    };
  } catch (e) {
    logError("fetchYahooPrice parse [" + symbol + "]", e);
    return null;
  }
}

// ========================================
// fetch Yahoo Finance ย้อนหลัง N วัน
// ใช้สำหรับคำนวณ ATR14, MA
// ========================================
function fetchYahooHistory(symbol, days) {
  days = days || 20;
  const range = days <= 30 ? "1mo" :
                days <= 90 ? "3mo" :
                days <= 180 ? "6mo" : "1y";

  const url  = "https://query1.finance.yahoo.com/v8/finance/chart/" +
    symbol + "?interval=1d&range=" + range;
  const opts = {
    muteHttpExceptions: true,
    headers: { "User-Agent": "Mozilla/5.0" }
  };

  const resp = fetchWithRetry(url, opts, 3);
  if (!resp) return null;

  try {
    const data   = JSON.parse(resp.getContentText());
    const result = data?.chart?.result?.[0];
    if (!result) return null;

    const timestamps = result.timestamp || [];
    const quote      = result.indicators?.quote?.[0] || {};

    return timestamps.map((ts, i) => ({
      date:   new Date(ts * 1000),
      open:   quote.open?.[i]   || 0,
      high:   quote.high?.[i]   || 0,
      low:    quote.low?.[i]    || 0,
      close:  quote.close?.[i]  || 0,
      volume: quote.volume?.[i] || 0
    })).filter(d => d.close > 0);

  } catch (e) {
    logError("fetchYahooHistory parse [" + symbol + "]", e);
    return null;
  }
}


// ========================================
// helper.gs — เพิ่ม logInfo และ logError ไปอยู่ webapp_03_analyzeใ.gs
// ========================================

function logInfo(fnName, message) {
  Logger.log("[INFO] " + fnName + " : " + message);
}




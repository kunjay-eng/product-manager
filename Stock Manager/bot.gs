// ========================================
// bot.gs
// ========================================


function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    if (!update.message) return;

    const text  = String(update.message.text || "").trim();
    const lower = text.toLowerCase();

    // Commands ที่มี parameter
    if (lower.startsWith("/nav "))      return cmdUpdateNAV(text);  //✅
    if (lower.startsWith("/atr14 "))    return cmdUpdateATR14(text);
    if (lower.startsWith("/exchange ")) return cmdExchange(text); //✅
    if (lower.startsWith("/atrfind ")) return  cmdATRFind(text);
    if (lower.startsWith("/backtest ")) return cmdBacktest(text);
    // ใน doPost() เพิ่มก่อน switch
    if (lower.startsWith("/risk ")) return cmdRiskReport(text);
    if (lower.startsWith("/analyze ")) return cmdAnalyze(text);
    if (lower.startsWith("/scan"))     return  cmdScanNow(text);   // ← เพิ่มบรรทัดนี้


    // TICKER PRICE BUDGET
    if (!text.startsWith("/") && _isTickerWithParams(text)) {
      const parts  = text.trim().split(/\s+/);
      const ticker = parts[0].toUpperCase();
      const price  = parseFloat(parts[1]);
      const budget = parseFloat(parts[2]);
      return sendStockInfo(ticker, price, budget);
    }

    // TICKER อย่างเดียว
    if (!text.startsWith("/") && _isTickerFormat(text)) {
      return sendStockInfo(text.toUpperCase());
    }

    
    
    switch (lower) {
  // 📊 Portfolio
      case "/dime":        sendSummaryDashboard();      break; // ✅ Dime! Summary
   //   case "/portfolio":   sendPortfolioReport(); break;  //  เปลี่ยนชื่อ
      case "/holdings":    sendHoldingsByProfit();      break; // ✅ Holding Report
   //   case "/weekly":     sendWeeklyPerformance(); break;
      case "/dashboard":  sendPortfolioDashboard(); break; //✅


  // 🔥ATR
  //   case "/atr":         sendATRAll();                break;
  //   case "/atralert":    sendATRAlert();              break;
      case "/atr":       sendATRDailySummary(); break;  // ฟังก์ชัน 1  //✅
  //    case "/atrstock":  sendATRByStock();      break;  // ฟังก์ชัน 2
  //    case "/atralert":  sendATRDetail();       break;  // ฟังก์ชัน 3  //✅
      case "/atrdash":   sendATRDashboard();    break;  // ฟังก์ชัน 4  //✅
      case "/atr14list": sendATR14List();      break;  //✅
      case "/atrfind":    // ❌ ไม่ทำงานเพราะต้องมี parameter  //✅
      sendTelegramSafe(
    "💡 วิธีใช้: /atrfind TICKER\n" +
    "เช่น /atrfind VOO"
                      );
    break;

      
    
      // 🏛️ กองทุน
      case "/navlist":      sendFundNAVList();           break; //✅
      //case "/nav":          cmdUpdateNAV();           break;  

      // 💰 Dividend
      case "/dividend":     sendMonthlyDividendReport(); break; //✅

      // 💱 FX
      case "/checkfx":      sendFX();                   break; //✅

      // 🔄 Update
      case "/updatepcs":   updatePrices(); break;  //  //✅ ดึงทั้ง TH+US
      case "/update":      updateSetting();             break; //  //✅ Update Setting รวม
      // alertlog
      case "/alertlog":     sendAlertLog();          break; //✅
      case "/alertclear":   clearAlertLog();         break; //✅


      // 🩺 System
      case "/health":      healthCheck();               break; //✅
      case "/help":
      case "/helpme":      showHelp();                  break; //✅

      default:
        sendTelegramSafe(
          "❓ ไม่พบคำสั่ง\n" +
          "พิมพ์ชื่อหุ้นตรงๆ เช่น MU, VOO, SCB.BK\n" +
          "หรือ /helpme เพื่อดูคำสั่งทั้งหมด"
        );
    }
  } catch (err) {
    sendTelegramSafe("❌ ERROR\n\n" + err.toString());
  }
}

function _isTickerWithParams(text) {
  return /^[A-Za-z0-9.\-]{1,12}\s+\d+(\.\d+)?\s+\d+(\.\d+)?$/.test(text.trim());
}

function _isTickerFormat(text) {
  return /^[A-Za-z0-9.\-]{1,12}$/.test(text.trim());
}




// ----------------------------------------
// เช็คว่าข้อความเป็นชื่อหุ้นหรือเปล่า
// รองรับ: TSM, VOO, SCB.BK, AAPL
// ----------------------------------------


// ตรวจ format: TICKER PRICE หรือ TICKER PRICE BUDGET
// (2 ตัวเลข แต่ตัวที่ 2 ใหญ่กว่า = budget ไม่ใช่จำนวนหุ้น)
function _isTickerBuyByBudget(text) {
  const parts = text.trim().split(/\s+/);
  if (parts.length < 2 || parts.length > 3) return false;
  if (!_isTickerFormat(parts[0])) return false;
  if (isNaN(parseFloat(parts[1]))) return false;
  // ถ้ามี 3 ส่วน ตรวจว่าเป็นตัวเลขด้วย
  if (parts.length === 3 && isNaN(parseFloat(parts[2]))) return false;
  return true;
}


// ----------------------------------------
// Webhook
// ----------------------------------------
// ========================================
// bot.gs — Telegram helpers
// ========================================




// ── Webhook helpers ──────────────────────
function setWebhook() {
  const { token } = _getTelegramCreds();
  const webAppUrl = "https://script.google.com/macros/s/AKfycby3Xu18oY97N9fJH9gCHI5MqNGo7ao4ElDVAVyJu3YcC9Ix16C64uz0T3fqwCHbPLGt/exec";
  const url = "https://api.telegram.org/bot" + token +
    "/setWebhook?url=" + encodeURIComponent(webAppUrl) +
    "&drop_pending_updates=true";
  Logger.log(UrlFetchApp.fetch(url).getContentText());
}

function deleteWebhook() {
  const { token } = _getTelegramCreds();
  const url = "https://api.telegram.org/bot" + token +
    "/deleteWebhook?drop_pending_updates=true";
  Logger.log(UrlFetchApp.fetch(url).getContentText());
}

function getWebhookInfo() {
  const { token } = _getTelegramCreds();
  const url = "https://api.telegram.org/bot" + token + "/getWebhookInfo";
  Logger.log(UrlFetchApp.fetch(url).getContentText());
}





// ----------------------------------------
// Setup Triggers ทั้งหมดในครั้งเดียว
// ----------------------------------------
function setupAllTriggers() {
  createAlertCheckTriggers();   // ← แก้จาก createAlertTriggers()
  createFXTrigger();
  createUpdateTrigger();
  createPortfolioTrigger();
  createATRTrigger();
  createWeeklyTrigger();
  createMonthlyDividendTrigger();
  createSummaryDashboardTrigger();
  createNAVReminderTrigger();
  createATR14ReminderTrigger();
  createRSITrigger();
  createDCAAlertTrigger();   // ✅ ข้อ 8
  Logger.log("✅ All Triggers Setup Done");
}


// ----------------------------------------
// ตรวจสอบ trigger ทั้งหมดที่ตั้งไว้
// ----------------------------------------
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log("Total triggers: " + triggers.length);


  triggers.forEach(t => {
    Logger.log(t.getHandlerFunction() + " | " + t.getEventType());
  });
}


// ----------------------------------------
// ลบ trigger ทั้งหมด (ใช้เวลาต้องการ reset)
// ----------------------------------------
function deleteAllTriggers() {
  ScriptApp.getProjectTriggers()
    .forEach(t => ScriptApp.deleteTrigger(t));


  Logger.log("🗑️ ลบ Trigger ทั้งหมดแล้ว");
}














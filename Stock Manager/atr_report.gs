// ========================================
// atr_report.gs
// 4 ฟังก์ชัน ATR Report
// ========================================

// ----------------------------------------
// ดึงข้อมูล ATR rows ทั้งหมด
// ----------------------------------------
function _getATRRows(sheetName, startRow) {
  const sheet   = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const numRows = lastRow - startRow + 1;
  const rows    = sheet.getRange(startRow, 1, numRows, 25).getValues(); // ✅ ขยายถึง col Y
  const result  = [];

  rows.forEach((row, i) => {
    const symbol = row[ATR_COL.SYMBOL];
    if (!symbol || symbol === "") return;

    result.push({
      symbol:         String(symbol),
      tradeStyle:     String(row[ATR_COL.TRADE_STYLE]   || ""),
      buyPrice:       Number(row[ATR_COL.BUY_PRICE])     || 0,
      highestClose:   Number(row[ATR_COL.HIGHEST_CLOSE]) || 0,
      atr:            Number(row[ATR_COL.ATR])            || 0,
      multiplier:     Number(row[ATR_COL.MULTIPLIER])    || 0,
      trailingStop:   Number(row[ATR_COL.TRAILING_STOP]) || 0,
      status:         String(row[ATR_COL.STATUS]         || ""),
      minProfit:      Number(row[ATR_COL.MIN_PROFIT])    || 0,
      shares:         Number(row[ATR_COL.SHARES])        || 0,
      priceNow:       Number(row[ATR_COL.PRICE_NOW])     || 0,
      plNow:          Number(row[ATR_COL.PL_NOW])        || 0,
      takeProfit:     Number(row[ATR_COL.TAKE_PROFIT])   || 0,
      rr:             Number(row[ATR_COL.RR])             || 0,
      profitTarget:   Number(row[ATR_COL.PROFIT_TARGET]) || 0,
      totalProfit:    Number(row[ATR_COL.TOTAL_PROFIT])  || 0,
      pctProfit:      Number(row[ATR_COL.PCT_PROFIT])    || 0,
      maxRisk:        Number(row[ATR_COL.MAX_RISK])      || 0,
      summaryStatus:  String(row[ATR_COL.SUMMARY_STATUS] || ""), // ✅ col W
      stopDistance:   Number(row[ATR_COL.STOP_DISTANCE]) || 0,  // ✅ col X
      riskStatus:     String(row[ATR_COL.RISK_STATUS]    || "")  // ✅ col Y
    });
  });

  return result;
}


// ══════════════════════════════════════════════════════════
// เขียนราคาสดกลับเข้าไปในชีต ATR Portfolio (US/TH)
// เรียกจากหน้า Stock Detail (Fast Analysis) ตอนกด ▶️ auto-refresh
// เขียนแค่คอลัมน์ PRICE_NOW เท่านั้น — คอลัมน์อื่น (plNow, stopDistance,
// riskStatus ฯลฯ) ปล่อยให้สูตร/trigger เดิมของชีตคำนวณต่อเอง ไม่เขียนทับ
// ══════════════════════════════════════════════════════════
// ❌ ลบทิ้ง — ห้ามเขียนทับเซลล์นี้เพราะเป็นสูตรดึงจาก Holdings sheet อยู่แล้ว

//function writeLivePriceToATRSheet(sheetName, startRow, ticker, livePrice) {
//  try {
//    const sheet = getSheet(sheetName);
//    const lastRow = sheet.getLastRow();
//    if (lastRow < startRow) return { success: false, error: 'ไม่มีข้อมูลในชีต' };

//    const numRows = lastRow - startRow + 1;
//    const rows = sheet.getRange(startRow, 1, numRows, 25).getValues();

//    for (let i = 0; i < rows.length; i++) {
//      const symbol = String(rows[i][ATR_COL.SYMBOL] || '').trim().toUpperCase();
//     if (symbol !== ticker.toUpperCase()) continue;

//     const sheetRow = startRow + i;
//     const priceCol = ATR_COL.PRICE_NOW + 1; // ATR_COL เป็น 0-based → แปลงเป็นเลขคอลัมน์จริง
//     sheet.getRange(sheetRow, priceCol).setValue(livePrice);
//    return { success: true };
//  }

//  return { success: false, error: 'ไม่พบ ' + ticker + ' ในชีต ATR Portfolio' };
//} catch (e) {
//    logError('writeLivePriceToATRSheet', e);
//   return { success: false, error: e.message };
// }
//}


function testSNDK() {
  Logger.log(JSON.stringify(getFastSignal('SNDK'), null, 2));
}


// ----------------------------------------
// Status emoji + สถานะสั้น
// ----------------------------------------
function _statusShort(status) {
  if (status.includes("Take Profit"))   return { emoji: "🟢", short: "Take Profit" };
  if (status.includes("รันเทรนด์"))    return { emoji: "✅", short: "รัน" };
  if (status.includes("Cut"))           return { emoji: "🟡", short: "Cut / ออน" };
  if (status.includes("ออน"))          return { emoji: "🟡", short: "Cut / ออน" };
  return { emoji: "⚪", short: status };
}

// ----------------------------------------
// Risk Status ตาม Stop Distance%
// ----------------------------------------
function _riskStatus(priceNow, trailingStop) {
  if (trailingStop <= 0 || priceNow <= 0) return "";
  const dist = ((priceNow - trailingStop) / trailingStop) * 100;
  if (dist > 10)  return "✅ ปลอดภัย";
  if (dist >= 5)  return "🟡 เฝ้าระวัง";
  return "🔴 ใกล้โดน Stop";
}

// ----------------------------------------
// ฟังก์ชัน 1: สรุปพอร์ตประจำวัน
// แยกหมวด รัน / Cut ออน / Take Profit
// /atr
// ----------------------------------------
function sendATRDailySummary() {
  try {
    
    // ✅ ควรเป็นแบบนี้

  const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
  const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);


    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy"
    );

    let msg =
      "📋 สรุปพอร์ตประจำวัน  |  " + today + "\n\n";

    const buildSection = (rows, flag) => {
      const running  = rows.filter(r => r.status.includes("รันเทรนด์")).map(r => r.symbol);
      const cut      = rows.filter(r => r.status.includes("Cut") || r.status.includes("ออน")).map(r => r.symbol);
      const tp       = rows.filter(r => r.status.includes("Take Profit")).map(r => r.symbol);
      const warnRR   = rows.filter(r => r.rr < 0).map(r => r.symbol);

      let s = "";
      if (running.length) s += "✅ รันตามแผน : \n " + running.join(", ") + "\n\n";
      if (cut.length)     s += "🟡 Cut : \n" + cut.join(", ")     + "\n\n";
      if (tp.length)      s += "🟢 Take Profit: \n" + tp.join(", ")     + "\n\n";
      if (warnRR.length)  s += "\n⚠️ ต้องระวัง (R/R ติดลบ)\n" +
                               flag + " " + warnRR.join(", ") + "\n";
      return s;
    };

    if (usRows.length > 0) {
      msg += "─────────────────\n";
      msg += "🇺🇸 หุ้นสหรัฐ\n\n";
     
      msg += buildSection(usRows, "🇺🇸");
    }

    if (thRows.length > 0) {
      msg += "─────────────────\n";  
      msg += "\n🇹🇭 หุ้นไทย\n\n";
      
      msg += buildSection(thRows, "🇹🇭");
    }

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendATRDailySummary", e);
  }
}

// ----------------------------------------
// ฟังก์ชัน 2: รายตัว ระบุสกุลเงินในการ์ด
// แสดงทุกตัว แยก US/TH
// /atr
// ----------------------------------------
function sendATRByStock() {
  try {
    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    // TH
    for (const r of thRows) {
      const st = _statusShort(r.status);
      const riskSt = _riskStatus(r.priceNow, r.trailingStop);
      const maxRiskPct = r.buyPrice > 0
        ? ((r.trailingStop - r.buyPrice) / r.buyPrice) * 100 : 0;

      let msg =
        "📊 " + r.symbol + " | " + r.tradeStyle + " 🇹🇭\n" +
        "─────────────\n" +
        "💰  ซื้อที่: ฿" + fmt(r.buyPrice) +
          " | สูงสุด: ฿" + fmt(r.highestClose) + "\n" +
        "🎯  TP: ฿" + fmt(r.takeProfit) +
          " | Stop: ฿" + fmt(r.trailingStop) + "\n" +
        "📉  Max Risk: " + signTHB(r.maxRisk) +
          " (" + fmt(maxRiskPct) + "%)\n" +
        "📈  R/R: " + fmt(r.rr) + "x\n" +
        st.emoji + "  สถานะ: " + st.short + "\n";

      if (riskSt) msg += "⚠️  Risk: " + riskSt + "\n";

      sendTelegramSafe(msg);
      Utilities.sleep(300);
    }

    // US
    for (const r of usRows) {
      const st = _statusShort(r.status);
      const riskSt = _riskStatus(r.priceNow, r.trailingStop);
      const maxRiskPct = r.buyPrice > 0
        ? ((r.trailingStop - r.buyPrice) / r.buyPrice) * 100 : 0;

      let msg =
        "📊 " + r.symbol + " | " + r.tradeStyle + " 🇺🇸\n" +
        "─────────────\n" +
        "💰  ซื้อที่: $" + fmt(r.buyPrice) +
          " | สูงสุด: $" + fmt(r.highestClose) + "\n" +
        "🎯  TP: $" + fmt(r.takeProfit) +
          " | Stop: $" + fmt(r.trailingStop) + "\n" +
        "📉  Max Risk: " + signUSD(r.maxRisk) +
          " (" + fmt(maxRiskPct) + "%)\n" +
        "📈  R/R: " + fmt(r.rr) + "x\n" +
        st.emoji + "  สถานะ: " + st.short + "\n";

      if (riskSt) msg += "⚠️  Risk: " + riskSt + "\n";

      sendTelegramSafe(msg);
      Utilities.sleep(300);
    }
  } catch (e) {
    sendTelegramError("sendATRByStock", e);
  }
}

// ----------------------------------------
// ฟังก์ชัน 3: ละเอียด Position Trade
// แสดงข้อมูลครบ รวม R/R, Max Risk, สรุป
// /atralert
// ----------------------------------------
function sendATRDetail() {
  try {
    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);
    const all = [
      ...thRows.map(r => ({ ...r, isTH: true  })),
      ...usRows.map(r => ({ ...r, isTH: false }))
    ];

    for (const r of all) {
      const cur  = r.isTH ? "฿" : "$";
      const flag = r.isTH ? "🇹🇭" : "🇺🇸";

      // ── P/L ──
      const plPct = r.buyPrice > 0 && r.priceNow > 0
        ? ((r.priceNow - r.buyPrice) / r.buyPrice) * 100 : 0;

      // ── Profits ──
      const currentProfit = (r.priceNow    - r.buyPrice) * r.shares;
      const lockedProfit  = (r.trailingStop - r.buyPrice) * r.shares;
      const peakProfit    = (r.highestClose - r.buyPrice) * r.shares; // ✅ Peak

      // ── Reward Locked % ──
      const rewardLocked = peakProfit > 0
        ? (lockedProfit / peakProfit) * 100 : 0;

      // ── Distance to Stop ──
      const distToStop = r.trailingStop > 0 && r.priceNow > 0
        ? ((r.priceNow - r.trailingStop) / r.priceNow) * 100 : 0;

      // emoji ตาม distance
      const distEmoji =
        distToStop > 10 ? "🟢" :
        distToStop >= 5 ? "🟡" : "🔴";

      // ── Status Logic ──
      const isAboveTP   = r.priceNow >= r.takeProfit;
      const isBelowStop = r.priceNow <= r.trailingStop;
      const isNearStop  = !isBelowStop && distToStop < 5;

      let mainStatus = "";
      if (isBelowStop) {
        mainStatus = "🚨 SELL SIGNAL\nราคาหลุด Trailing Stop";
      } else if (isAboveTP) {
        mainStatus = "🎯 TARGET ACHIEVED\n🟢 HOLD TREND";
      } else if (isNearStop) {
        mainStatus = "⚠️ WATCH — ใกล้ Stop\n🔄 กำลังรันเทรนด์";
      } else {
        mainStatus = "🔄 กำลังรันเทรนด์";
      }

      let msg =
        "🔍 " + r.symbol + " | " + r.tradeStyle + " " + flag + "\n" +
        "─────────────────\n\n" +

        "💰 ราคาซื้อ       : " + cur + fmt(r.buyPrice)    + "\n" +
        "📈 Highest Close  : " + cur + fmt(r.highestClose) + "\n" +
        "📊 ราคาปัจจุบัน  : " + cur + fmt(r.priceNow)     + "\n" +
        plEmoji(plPct) +
          " P/L ปัจจุบัน  : " + signPct(plPct)            + "\n\n" +

        "⚙️ ATR(14)        : " + fmt(r.atr)        + "\n" +
        "⚙️ Multiplier     : " + r.multiplier       + "\n\n" +

        "📍 Stop & Target\n" +
        "🛑 Trailing Stop  : " + cur + fmt(r.trailingStop) + "\n" +
        "🎯 Initial Target : " + cur + fmt(r.takeProfit)   +
          (isAboveTP ? " ✅ Achieved" : "") + "\n" +
        "⚖️ Risk/Reward    : " + fmt(r.rr) + "x\n\n" +

        "━━━━━━━━━━━━\n\n" +
        "💹 ผลลัพธ์\n\n" +
        "💰 Current Profit : " + signStr(cur, currentProfit) + "\n" +
        "🔒 Locked Profit  : " + signStr(cur, lockedProfit)  + "\n" +
        "📈 Peak Profit    : " + signStr(cur, peakProfit)    + "\n" +  // ✅ เปลี่ยนชื่อ
        "📊 Reward Locked  : " + fmt(rewardLocked) + "%\n" +           // ✅ เพิ่มใหม่
        "📦 จำนวนหุ้น      : " + fmt(r.shares) + " หุ้น\n\n" +

        "━━━━━━━━━━━━\n\n" +
        "📌 สถานะปัจจุบัน\n\n" +
        mainStatus + "\n\n" +

        // ✅ Distance to Stop
        distEmoji + " ราคาอยู่เหนือ Stop " + fmt(distToStop) + "%\n" +
        "🛑 Stop ล่าสุด : " + cur + fmt(r.trailingStop) + "\n" +
        "\n━━━━━━━━━━━━";

      sendTelegramSafe(msg);
      Utilities.sleep(300);
    }
  } catch (e) {
    sendTelegramError("sendATRDetail", e);
  }
}


// ── Helper: format signed value ──
function signStr(cur, val) {
  if (val >= 0) return "+" + cur + fmt(val);
  return "-" + cur + fmt(Math.abs(val));
}


// ----------------------------------------
// ฟังก์ชัน 4: Dashboard สั้น จัดคอลัมน์
// /atrdash
// ----------------------------------------
function sendATRDashboard() {
  try {
    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy"
    );

    let msg = "📋 พอร์ตรวม  |  " + today + "\n\n";

    const styleShort = s => {
      if (s.includes("Long-term")) return "LT Invest";
      if (s.includes("Position"))  return "Position";
      if (s.includes("Swing"))     return "Swing";
      return s.substring(0, 8);
    };

    const formatRow = (flag, r) => {
      const st   = _statusShort(r.status);
      const rrStr = (r.rr >= 0 ? "+" : "") + fmt(r.rr) + "x";
      return flag + " " +
        r.symbol.padEnd(7) + " " +
       // styleShort(r.tradeStyle).padEnd(10) + " " +
        rrStr.padStart(7) + "  " +
        st.emoji + " " + st.short + "\n";
    };

    // Header
    msg += "  " +
      "Symbol ".padEnd(8) +
    //  "Style     ".padEnd(11) +
      "R/R    " +
      "Status\n";
    msg += "─────────────────\n";

    // US
    if (usRows.length > 0) {
      usRows.forEach(r => { msg += formatRow("🇺🇸", r); });
    }

    if (usRows.length > 0 && thRows.length > 0) {
      msg += "─────────────────\n";
    }

    // TH
    if (thRows.length > 0) {
      thRows.forEach(r => { msg += formatRow("🇹🇭", r); });
    }

    // ต้องดู (R/R ติดลบ)
    const allRows = [...usRows, ...thRows];
    const warnList = allRows.filter(r => r.rr < 0);
    if (warnList.length > 0) {
      msg += "─────────────────────────────\n";
      msg += "🔴 ต้องดู: " +
        warnList.map(r => r.symbol).join(", ");
    }

    sendTelegramSafe(msg);
  } catch (e) {
    sendTelegramError("sendATRDashboard", e);
  }
}




// ----------------------------------------
// ค้นหาหุ้นใน ATR Portfolio
// พิมพ์: /atrfind VOO หรือ /atrfind SCB
// ----------------------------------------
function cmdATRFind(text) {
  try {
    const parts  = text.trim().split(/\s+/);
    const ticker = parts[1] ? parts[1].toUpperCase().trim() : "";

    if (!ticker) {
      sendTelegramSafe(
        "❌ กรุณาระบุชื่อหุ้น\n\n" +
        "✅ วิธีใช้: /atrfind TICKER\n" +
        "📌 เช่น /atrfind VOO หรือ /atrfind SCB"
      );
      return;
    }

    // ค้นหาใน US และ TH
    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    let found = usRows.find(r => r.symbol.toUpperCase() === ticker);
    let isTH  = false;

    if (!found) {
      found = thRows.find(r => r.symbol.toUpperCase() === ticker);
      if (found) isTH = true;
    }

    if (!found) {
      const allSymbols = [
        ...usRows.map(r => "🇺🇸 " + r.symbol),
        ...thRows.map(r => "🇹🇭 " + r.symbol)
      ];
      sendTelegramSafe(
        "❌ ไม่พบ: " + ticker + "\n\n" +
        "📋 หุ้นทั้งหมดในระบบ:\n" +
        allSymbols.join("\n")
      );
      return;
    }

    // ✅ ใช้ format เดียวกับ sendATRDetail()
    const r   = { ...found, isTH };
    const cur = isTH ? "฿" : "$";
    const flag= isTH ? "🇹🇭" : "🇺🇸";

    const plPct = r.buyPrice > 0 && r.priceNow > 0
      ? ((r.priceNow - r.buyPrice) / r.buyPrice) * 100 : 0;

    const currentProfit = (r.priceNow     - r.buyPrice) * r.shares;
    const lockedProfit  = (r.trailingStop - r.buyPrice) * r.shares;
    const peakProfit    = (r.highestClose - r.buyPrice) * r.shares;
    const rewardLocked  = peakProfit > 0
      ? (lockedProfit / peakProfit) * 100 : 0;

    const distToStop = r.trailingStop > 0 && r.priceNow > 0
      ? ((r.priceNow - r.trailingStop) / r.priceNow) * 100 : 0;

    const distEmoji =
      distToStop > 10 ? "🟢" :
      distToStop >= 5 ? "🟡" : "🔴";

    const isAboveTP   = r.priceNow >= r.takeProfit;
    const isBelowStop = r.priceNow <= r.trailingStop;
    const isNearStop  = !isBelowStop && distToStop < 5;

    let mainStatus = "";
    if (isBelowStop) {
      mainStatus = "🚨 SELL SIGNAL\nราคาหลุด Trailing Stop";
    } else if (isAboveTP) {
      mainStatus = "🎯 TARGET ACHIEVED\n🟢 HOLD TREND";
    } else if (isNearStop) {
      mainStatus = "⚠️ WATCH — ใกล้ Stop\n🔄 กำลังรันเทรนด์";
    } else {
      mainStatus = "🔄 กำลังรันเทรนด์";
    }

    const msg =
      "🔍 " + r.symbol + " | " + r.tradeStyle + " " + flag + "\n" +
      "─────────────────\n\n" +

      "💰 ราคาซื้อ       : " + cur + fmt(r.buyPrice)    + "\n" +
      "📈 Highest Close  : " + cur + fmt(r.highestClose) + "\n" +
      "📊 ราคาปัจจุบัน  : " + cur + fmt(r.priceNow)     + "\n" +
      plEmoji(plPct) +
        " P/L ปัจจุบัน  : " + signPct(plPct)            + "\n\n" +

      "⚙️ ATR(14)        : " + fmt(r.atr)   + "\n" +
      "⚙️ Multiplier     : " + r.multiplier + "\n\n" +

      "📍 Stop & Target\n" +
      "🛑 Trailing Stop  : " + cur + fmt(r.trailingStop) + "\n" +
      "🎯 Initial Target : " + cur + fmt(r.takeProfit)   +
        (isAboveTP ? " ✅ Achieved" : "") + "\n" +
      "⚖️ Risk/Reward    : " + fmt(r.rr) + "x\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "💹 ผลลัพธ์\n\n" +
      "💰 Current Profit : " + signStr(cur, currentProfit) + "\n" +
      "🔒 Locked Profit  : " + signStr(cur, lockedProfit)  + "\n" +
      "📈 Peak Profit    : " + signStr(cur, peakProfit)    + "\n" +
      "📊 Reward Locked  : " + fmt(rewardLocked) + "%\n"  +
      "📦 จำนวนหุ้น      : " + fmt(r.shares) + " หุ้น\n\n" +

      "━━━━━━━━━━━━\n\n" +
      "📌 สถานะปัจจุบัน\n\n" +
      mainStatus + "\n\n" +
      distEmoji + " ราคาอยู่เหนือ Stop +" + fmt(distToStop) + "%\n" +
      "🛑 Stop ล่าสุด : " + cur + fmt(r.trailingStop) + "\n" +
      "\n━━━━━━━━━━━━";

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("cmdATRFind", e);
  }
}

// ----------------------------------------
// แจ้งเตือน trigger — หุ้นที่ col W มี "ขายที่"
// Auto trigger ทุก 30 นาที
// ----------------------------------------
function checkATRTriggerAuto() {
  try {
    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    const triggered = [
      ...usRows.filter(r => r.summaryStatus.includes("ขายที่"))
               .map(r => ({ ...r, isTH: false })),
      ...thRows.filter(r => r.summaryStatus.includes("ขายที่"))
               .map(r => ({ ...r, isTH: true  }))
    ];

    if (triggered.length === 0) return;

    let msg =
      "🚨 ATR TRIGGER ALERT!\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "พบหุ้นที่ถึงจุดขาย " + triggered.length + " ตัว\n\n";

    for (const r of triggered) {
      const cur  = r.isTH ? "฿" : "$";
      const flag = r.isTH ? "🇹🇭" : "🇺🇸";

      msg +=
        flag + " " + r.symbol + " — " + r.tradeStyle + "\n" +
        "  📝 " + r.summaryStatus + "\n" +
        "  🛑 Stop : " + cur + fmt(r.trailingStop) + "\n" +
        "  🎯 TP   : " + cur + fmt(r.takeProfit)   + "\n\n";
    }

    sendTelegramSafe(msg);

  } catch (e) {
    logError("checkATRTriggerAuto", e);
  }
}



// ========================================
// /analyze TICKER — วิเคราะห์หุ้นที่ถืออยู่
// ดึงข้อมูลจาก ATR sheet + Yahoo Finance
// ========================================
function cmdAnalyze(text) {
  try {
    const parts  = text.trim().split(/\s+/);
    const ticker = parts[1] ? parts[1].toUpperCase() : "";

    if (!ticker) {
      sendTelegramSafe(
        "❌ กรุณาระบุชื่อหุ้น\n\n" +
        "✅ วิธีใช้: /analyze TICKER\n" +
        "📌 เช่น /analyze MRVL หรือ /analyze SCB"
      );
      return;
    }

    sendTelegramSafe("⏳ กำลังวิเคราะห์ " + ticker + "...");

    // ── ดึงจาก ATR sheet ──
    const usRows = _getATRRows(ATR_SHEETS.US, ATR_START_ROW.US);
    const thRows = _getATRRows(ATR_SHEETS.TH, ATR_START_ROW.TH);

    let atr  = usRows.find(r => r.symbol.toUpperCase() === ticker);
    let isTH = false;
    if (!atr) {
      atr  = thRows.find(r => r.symbol.toUpperCase() === ticker);
      isTH = !!atr;
    }

    if (!atr) {
      sendTelegramSafe(
        "❌ ไม่พบ " + ticker + " ใน ATR Portfolio\n\n" +
        "หุ้ที่มีในระบบ:\n" +
        [...usRows.map(r => "🇺🇸 " + r.symbol),
         ...thRows.map(r => "🇹🇭 " + r.symbol)].join("\n")
      );
      return;
    }

    // ── ดึงจาก Yahoo Finance ──
    const yahooSymbol = isTH ? ticker + ".BK" : ticker;
    const stockData   = _fetchStockData(yahooSymbol);
    if (!stockData) {
      sendTelegramSafe("❌ ดึงข้อมูล Yahoo ไม่ได้: " + ticker);
      return;
    }

    const cur    = isTH ? "฿" : "$";
    const flag   = isTH ? "🇹🇭" : "🇺🇸";
    const price  = atr.priceNow > 0 ? atr.priceNow : stockData.price;
    const ma20   = stockData.ma20;
    const ma50   = stockData.ma50;
    const ma200  = stockData.ma200;
    const rsi    = stockData.rsi;
    const volNow = stockData.volNow;
    const volAvg = stockData.volAvg20;
    const volRatio = stockData.volRatio;

    // ── คำนวณ P/L ──
    const plPct = atr.buyPrice > 0
      ? ((price - atr.buyPrice) / atr.buyPrice) * 100 : 0;

    // ── 1. Risk Management — Trailing Stop ──
    const belowStop  = price < atr.trailingStop;
    const distToStop = atr.trailingStop > 0
      ? ((price - atr.trailingStop) / price) * 100 : 0;

    // ── 2. P/L Assessment ──
    let plSignal = "";
    if (plPct > 20)        plSignal = "🟢 กำไรดี — เริ่มปกป้องกำไร";
    else if (plPct > 0)    plSignal = "🟡 กำไรน้อย — ดูแนวโน้มต่อ";
    else if (plPct > -10)  plSignal = "🟡 ติดลบเล็กน้อย — เฝ้าระวัง";
    else if (plPct > -15)  plSignal = "🔴 ติดลบ >10% — ระวัง";
    else                   plSignal = "🔴 ติดลบ >15% — พิจารณาตัด";

    // ── 3. Trend ──
    let trendSignal = "";
    if (price > ma20 && ma20 > ma50 && ma50 > ma200) {
      trendSignal = "🟢 ขาขึ้นแข็งแรง (Price>MA20>MA50>MA200)";
    } else if (price > ma50 && ma50 > ma200) {
      trendSignal = "🟡 ขาขึ้น แต่ราคาต่ำกว่า MA20";
    } else if (price < ma20 && price < ma50) {
      trendSignal = "🔴 ขาลง (ราคาต่ำกว่า MA20 และ MA50)";
    } else {
      trendSignal = "🟡 แนวโน้มผสม";
    }

    // ── 4. RSI ──
    let rsiSignal = "";
    if (rsi > 70)      rsiSignal = "🔴 RSI " + fmt(rsi) + " — ร้อนแรงเกินไป";
    else if (rsi > 50) rsiSignal = "🟢 RSI " + fmt(rsi) + " — แข็งแรง";
    else if (rsi > 30) rsiSignal = "🟡 RSI " + fmt(rsi) + " — อ่อนตัว";
    else               rsiSignal = "💡 RSI " + fmt(rsi) + " — Oversold";

    // ── 5. Volume ──
    let volSignal = "";
    if (plPct < 0 && volRatio > 1.5) {
      volSignal = "🔴 ราคาลง + Volume สูง ×" + fmt(volRatio) + " — สัญญาณลบ";
    } else if (plPct > 0 && volRatio > 1.5) {
      volSignal = "🟢 ราคาขึ้น + Volume สูง ×" + fmt(volRatio) + " — สัญญาณบวก";
    } else {
      volSignal = "🟡 Volume ปกติ ×" + fmt(volRatio);
    }

    // ── สรุป Decision ──
    const reasons  = [];
    const warnings = [];
    let   decision = "";
    let   decEmoji = "";
    let   plan     = "";

    if (belowStop) {
      // กฎข้อ 1 — ขายทันที
      decision = "ขายทันที";
      decEmoji = "🚨";
      reasons.push("❌ ราคาต่ำกว่า Trailing Stop " + cur + fmt(atr.trailingStop));
      reasons.push("   ระบบบอกว่าแนวโน้มเสียแล้ว");
      plan = "⛔ ขายออกทั้งหมดทันที\n" +
             "   Stop: " + cur + fmt(atr.trailingStop) + "\n" +
             "   ราคาปัจจุบัน: " + cur + fmt(price);

    } else if (plPct < -15) {
      // กฎข้อ 2 — พิจารณาตัดขาดทุน
      decision = "พิจารณาตัดขาดทุน";
      decEmoji = "🔴";
      reasons.push("❌ ขาดทุน " + fmt(plPct) + "% เกิน 15%");
      if (price < ma50) reasons.push("❌ ราคาต่ำกว่า MA50");
      plan = "⚠️ พิจารณาตัดขาดทุน\n" +
             "   ถ้าราคาไม่ฟื้นกลับเหนือ " + cur + fmt(ma50) + " (MA50)\n" +
             "   ให้ขายออกเพื่อจำกัดความเสียหาย";

    } else if (price > ma20 && rsi < 70 && !belowStop && plPct >= 0) {
      // ถือต่อ — แนวโน้มดี
      decision = "ถือ";
      decEmoji = "🟢";
      if (price > ma20)  reasons.push("✅ ราคายืนเหนือ MA20");
      if (ma20 > ma50)   reasons.push("✅ MA20 > MA50");
      if (ma50 > ma200)  reasons.push("✅ MA50 > MA200");
      if (volRatio > 1.5 && plPct > 0)
                         reasons.push("✅ Volume สนับสนุนการขึ้น");
      if (rsi > 50 && rsi < 70)
                         reasons.push("✅ RSI แข็งแรง");
      if (rsi >= 65)     warnings.push("⚠️ RSI เริ่มสูง ติดตามใกล้ชิด");
      plan = "📌 ถือรันกำไรต่อ\n" +
             "   หากราคาปิดต่ำกว่า " + cur + fmt(atr.trailingStop) +
             " ให้ขายทันที";

    } else if (plPct < 0 && price > ma50) {
      // ถือ — ยังอยู่เหนือ MA50
      decision = "ถือ (เฝ้าระวัง)";
      decEmoji = "🟡";
      reasons.push("✅ ราคายังอยู่เหนือ MA50");
      reasons.push("⚠️ ราคาต่ำกว่า MA20 — แรงขายระยะสั้น");
      if (rsi < 40) reasons.push("💡 RSI ต่ำ อาจดีดกลับ");
      plan = "👀 เฝ้าระวัง\n" +
             "   รอราคากลับขึ้นเหนือ MA20 (" + cur + fmt(ma20) + ")\n" +
             "   Stop: " + cur + fmt(atr.trailingStop);

    } else {
      decision = "เฝ้าระวัง";
      decEmoji = "🟡";
      reasons.push("🟡 สัญญาณไม่ชัดเจน");
      plan = "👀 ติดตามใกล้ชิด\n" +
             "   Stop: " + cur + fmt(atr.trailingStop);
    }

    // ── Build Message ──
    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm"
    );

    let msg =
      "📋 วิเคราะห์หุ้น " + flag + "\n" +
      "━━━━━━━━━━━━\n\n" +

      "🔍 Ticker      : " + atr.symbol      + "\n" +
      "📌 Strategy    : " + atr.tradeStyle  + "\n\n" +

      "💰 Entry Price : " + cur + fmt(atr.buyPrice)     + "\n" +
      "📈 Highest Close: " + cur + fmt(atr.highestClose) + "\n" +
      "📊 ราคาปัจจุบัน: " + cur + fmt(price)            + "\n\n" +

      "⚙️ ATR(14)     : " + fmt(atr.atr)        + "\n" +
      "⚙️ Multiplier  : " + atr.multiplier       + "\n\n" +

      "📉 Trailing Stop: " + cur + fmt(atr.trailingStop) + "\n" +
      "   ห่าง Stop   : " +
        (belowStop
          ? "🔴 ต่ำกว่า Stop " + fmt(Math.abs(distToStop)) + "%"
          : "+" + fmt(distToStop) + "%") + "\n\n" +

      "📊 P/L ปัจจุบัน: " + signPct(plPct) + "\n" +
      "   " + plSignal + "\n\n" +

      "━━━━━━━━━━━━\n" +
      "📈 Technical\n\n" +
      "   MA20  : " + cur + fmt(ma20)  + "\n" +
      "   MA50  : " + cur + fmt(ma50)  + "\n" +
      "   MA200 : " + cur + fmt(ma200) + "\n" +
      "   " + trendSignal + "\n\n" +
      "   " + rsiSignal   + "\n\n" +
      "📦 Volume    : " + _fmtVol(volNow) +
        " (avg20: " + _fmtVol(volAvg) + ")\n" +
      "   " + volSignal   + "\n\n" +

      "━━━━━━━━━━━━\n" +
      decEmoji + " คำแนะนำ : " + decision + "\n\n" +

      "เหตุผล:\n";

    reasons.forEach(r  => { msg += r  + "\n"; });
    warnings.forEach(w => { msg += w  + "\n"; });

    msg +=
      "\nแผน:\n" + plan + "\n\n" +
      "🕐 " + today;

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("cmdAnalyze", e);
  }
}




// ========================================
// dividend.gs
// /dividend
// ========================================

const MONTH_NAMES = [
  "","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม",
  "มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม",
  "พฤศจิกายน","ธันวาคม"
];

// ========================================
// dividend.gs
// /dividend — Monthly Dividend Report
// ========================================

function sendMonthlyDividendReport() {
  try {
    const now   = new Date();
    const year  = now.getFullYear();
    const month = now.getMonth() + 1;

    // ✅ ดึงปันผลเดือนนี้
    const monthly = getDividendMonthly(year, month);
    const ytd     = getDividendYTD();

    const monthName = Utilities.formatDate(now, "Asia/Bangkok", "MMMM yyyy");

    let msg =
      "💰 DIVIDEND REPORT\n" +
      "🕐 " + getNow() + "\n" +
      "━━━━━━━━━━━━\n\n" +
      "📅 " + monthName + "\n\n";

    // เดือนนี้
    if (!monthly || monthly.length === 0) {
      msg += "📭 ไม่มีปันผลเดือนนี้\n\n";
    } else {
      msg += "📋 รายการปันผลเดือนนี้\n";
      let totalTHB = 0;

      monthly.forEach(d => {
        const payDate = d.payDate instanceof Date
          ? Utilities.formatDate(d.payDate, "Asia/Bangkok", "dd/MM/yyyy")
          : String(d.payDate || "");

        const amtStr = d.currency === "USD"
          ? "$" + fmt(d.amt) + " (฿" + fmt(d.netTHB) + ")"
          : "฿" + fmt(d.netTHB);

        msg +=
          "  • " + d.ticker + " — " + d.company + "\n" +
          "    📅 " + payDate + "\n" +
          "    💰 " + amtStr + "\n\n";

        totalTHB += d.netTHB;
      });

      msg += "💵 รวมเดือนนี้ : ฿" + fmt(totalTHB) + "\n\n";
    }

    // YTD
    msg +=
      "━━━━━━━━━━━━\n" +
      "📊 YTD Summary\n" +
      "  🇹🇭 ปันผลไทย   : ฿" + fmt(ytd.thTHB)    + "\n" +
      "  🇺🇸 ปันผลสหรัฐ : $" + fmt(ytd.usUSD)    + "\n" +
      "  💰 รวม YTD     : ฿" + fmt(ytd.totalTHB) + "\n" +
      "  🎯 เป้าหมาย    : ฿" + fmt(ytd.target)   + "\n\n";

    // Progress
    const progress = ytd.target > 0
      ? Math.min((ytd.totalTHB / ytd.target) * 100, 100) : 0;
    const bar = _buildProgressBar(progress, 10);
    msg += bar + " " + fmt(progress) + "%";

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("sendMonthlyDividendReport", e);
  }
}

// ----------------------------------------
// Trigger Setup — ทุกวันที่ 1 ของเดือน 09:00
// ----------------------------------------
function createMonthlyDividendTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendMonthlyDividendReport")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // #6 Monthly Dividend — วันที่ 1 ของเดือน 09:00
  ScriptApp.newTrigger("sendMonthlyDividendReport")
    .timeBased().onMonthDay(1).atHour(9)
    .inTimezone("Asia/Bangkok").create();

  Logger.log("✅ Dividend Trigger: วันที่ 1 ของเดือน 09:00");

}











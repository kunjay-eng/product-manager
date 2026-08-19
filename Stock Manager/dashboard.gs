// ========================================
// dashboard.gs
// /dashboard — Portfolio Dashboard
// ========================================

function sendPortfolioDashboard() {
  try {
    sendTelegramSafe("⏳ กำลังสร้าง Portfolio Dashboard...");

    const d      = collectAllData();
    const master = getMasterData();
    const benches= getBenchmarks();
    const fxRate = d.fxRate;

    // ── Asset Allocation ──
    const totalAsset = d.totalPortTHB + d.cashTotal;
    const thPct      = totalAsset > 0 ? (d.thPortTHB   / totalAsset) * 100 : 0;
    const usPct      = totalAsset > 0 ? ((d.usPortUSD * fxRate) / totalAsset) * 100 : 0;
    const fundPct    = totalAsset > 0 ? (d.fundPortTHB / totalAsset) * 100 : 0;
    const cashPct    = totalAsset > 0 ? (d.cashTotal   / totalAsset) * 100 : 0;

    // ── Portfolio Return ──
    const totalReturn = d.totalUnrealTHB + d.totalRealTHB + d.divTotalTHB;
    const netInvest   = d.netInvest > 0 ? d.netInvest : master.initialInvest;
    const portRetPct  = netInvest > 0 ? (totalReturn / netInvest) * 100 : 0;

    // ── Benchmark ──
    // หา benchmark ที่ใกล้เคียงที่สุด (S&P500 สำหรับ portfolio หลัก)
    const sp500 = benches.find(b => b.name === "S&P500");
    const benchRetPct = sp500 ? sp500.return * 100 : 0;
    const alpha = portRetPct - benchRetPct;

    // ── Top/Worst Movers ──
    const thSorted = [...d.thHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);
    const usSorted = [...d.usHoldings].sort(
      (a, b) => b.unrealizedPct - a.unrealizedPct);

    // ── Financial Goal ──
    const portTarget  = master.portTarget;
    const progress    = portTarget > 0
      ? (totalAsset / portTarget) * 100 : 0;
    const progressBar = _buildProgressBar(progress, 10);

    // ── CAGR ──
    const startDate = master.startDate instanceof Date
      ? master.startDate : new Date(master.startDate);
    const daysDiff  = Math.max(1,
      (new Date() - startDate) / (1000 * 60 * 60 * 24));
    const years     = daysDiff / 365;
    const cagr      = netInvest > 0 && years > 0
      ? (Math.pow(totalAsset / netInvest, 1 / years) - 1) * 100 : 0;

    const today = Utilities.formatDate(
      new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm"
    );

    const medals = ["🥇", "🥈", "🥉"];

    let msg =
      "📊 PORTFOLIO DASHBOARD\n" +
      "🕐 " + today + "\n" +
      "━━━━━━━━━━━━━━━━━━\n\n" +

      "📊 Asset Allocation\n" +
      "🇹🇭 หุ้นไทย      : " + fmt(thPct)   + "%\n" +
      "🇺🇸 หุ้นสหรัฐฯ   : " + fmt(usPct)   + "%\n" +
      "🏛️ กองทุน        : " + fmt(fundPct) + "%\n" +
      "💵 เงินสด         : " + fmt(cashPct) + "%\n\n" +

      "━━━━━━━━━━━━━━━━━━\n\n" +

      "📈 Benchmark Comparison\n\n" +
      "Portfolio      : " + signPct(portRetPct) + "\n\n";

    // แสดง benchmark ทั้งหมด
    benches.forEach(b => {
      const retPct = b.return * 100;
      const alp    = portRetPct - retPct;
      msg +=
        b.asset.padEnd(14) + " vs " + b.name + " : " + signPct(retPct) + "\n" +
        "  " + (alp >= 0 ? "🏆" : "📉") +
        " Alpha : " + signPct(alp) + "\n\n";
    });

    msg += "\n━━━━━━━━━━━━━━━━━━\n🏆🇹🇭 Top Movers\n";

    // TH Ranking
    msg += "\n";
    thSorted.slice(0, 3).forEach((h, i) => {
      msg += medals[i] + " " + h.ticker.padEnd(8) +
        signPct(h.unrealizedPct * 100) + "\n";
    });
    const thWorst = thSorted.slice(-3).reverse();
    thWorst.forEach(h => {
      msg += "📉 " + h.ticker.padEnd(8) +
        signPct(h.unrealizedPct * 100) + "\n";
    });

    msg += "\n━━━━━━━━━━━━━━━━━━\n";

    // US Ranking
    msg += "🏆🇺🇸 Top Movers\n\n";
    usSorted.slice(0, 3).forEach((h, i) => {
      msg += medals[i] + " " + h.ticker.padEnd(8) +
        signPct(h.unrealizedPct * 100) + "\n";
    });
    const usWorst = usSorted.slice(-3).reverse();
    usWorst.forEach(h => {
      msg += "📉 " + h.ticker.padEnd(8) +
        signPct(h.unrealizedPct * 100) + "\n";
    });

    msg +=
      "\n━━━━━━━━━━━━━━━━━━\n\n" +

      "🎯 Financial Goal\n" +
      "ลงทุน            : " + fmtTHB(netInvest)   + "\n" +
      "เป้าหมายพอร์ต   : " + fmtTHB(portTarget)  + "\n" +
      "ปัจจุบัน         : " + fmtTHB(totalAsset)  + "\n" +
      "Progress         : " + fmt(progress) + "%\n\n" +
      progressBar + "\n\n" +

      "━━━━━━━━━━━━━━━━━━\n\n" +

      "🚀 CAGR\n" +
      "เริ่มลงทุน : " + fmt(years) + " Years\n" +
      "CAGR           : " + signPct(cagr) + "/ปี";

    sendTelegramSafe(msg);

  } catch (e) {
    sendTelegramError("sendPortfolioDashboard", e);
  }
}

// Progress Bar
function _buildProgressBar(pct, total) {
  const filled = Math.min(Math.round((pct / 100) * total), total);
  const empty  = total - filled;
  return "🟩".repeat(filled) + "⬜".repeat(empty);
}

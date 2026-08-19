// ============================================================
// webapp_15_portfolio_history.gs — สร้างใหม่ทั้งหมด
// ============================================================

const PORTFOLIO_LOG_SHEET_NAME = 'Portfolio_Value_Log';
const GFINANCE_SCRATCH_SHEET = '_GFinance_Scratch';

const BENCHMARK_SYMBOLS = {
  SET: 'SET',
  SP500: 'SPY'
};

function _ensurePortfolioLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PORTFOLIO_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTFOLIO_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Date', 'TotalAssetTHB']]);
  }
  return sheet;
}

function logPortfolioValueSnapshot() {
  try {
    const summary = getSummaryData();
    if (!summary || !summary.success) {
      logError('logPortfolioValueSnapshot', new Error('getSummaryData ล้มเหลว'));
      return;
    }
    const sheet = _ensurePortfolioLogSheet();
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r =>
        r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[0])
      );
      const idx = dates.indexOf(today);
      if (idx !== -1) { sheet.getRange(idx + 2, 2).setValue(summary.totalAssetTHB); return; }
    }
    sheet.appendRow([new Date(), summary.totalAssetTHB]);
  } catch (e) {
    logError('logPortfolioValueSnapshot', e);
  }
}

function createPortfolioSnapshotTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'logPortfolioValueSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('logPortfolioValueSnapshot').timeBased().everyDays(1).atHour(23).create();
  logPortfolioValueSnapshot();
}

// ── เรียกจากปุ่ม "เริ่มเก็บข้อมูลตอนนี้" ในหน้า UI — ตั้ง trigger + สร้างจุดข้อมูลแรกทันที ──
function initPortfolioHistoryTracking() {
  try {
    createPortfolioSnapshotTrigger();
    return { success: true };
  } catch (e) {
    logError('initPortfolioHistoryTracking', e);
    return { success: false, error: e.message };
  }
}

function getPortfolioValueHistory() {
  try {
    const sheet = _ensurePortfolioLogSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, points: [], note: 'ยังไม่มีข้อมูลสะสม' };

    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues()
      .filter(r => r[0] && r[1])
      .sort((a, b) => new Date(a[0]) - new Date(b[0]))
      .map(r => ({
        date: Utilities.formatDate(new Date(r[0]), 'Asia/Bangkok', 'dd/MM/yy'),
        value: parseFloat(r[1])
      }));

    return { success: true, points: rows };
  } catch (e) {
    logError('getPortfolioValueHistory', e);
    return { success: false, error: e.message };
  }
}

// ── ดึงราคาปิดในช่วง 7 วันถัดจาก startDate แล้วเอาค่าแรกที่ใช้ได้ — กันวันหยุด/เสาร์-อาทิตย์ที่ทำให้เป็น N/A ──
function _getGoogleFinanceCloseNear(symbol, dateObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let scratch = ss.getSheetByName(GFINANCE_SCRATCH_SHEET);
  if (!scratch) { scratch = ss.insertSheet(GFINANCE_SCRATCH_SHEET); scratch.hideSheet(); }
  scratch.getRange('A1:J20').clearContent();

  const endDate = new Date(dateObj.getTime() + 7 * 24 * 60 * 60 * 1000);
  const fmt = d => `DATE(${d.getFullYear()},${d.getMonth()+1},${d.getDate()})`;
  scratch.getRange('A1').setFormula(
    `=GOOGLEFINANCE("${symbol}","close",${fmt(dateObj)},${fmt(endDate)})`
  );
  SpreadsheetApp.flush();

  // ── retry สั้นๆ กันกรณี GOOGLEFINANCE ยังไม่ evaluate เสร็จ ──
  let values = null;
  for (let i = 0; i < 6; i++) {
    Utilities.sleep(500);
    values = scratch.getRange('A1:B10').getValues();
    const hasData = values.some(r => isFinite(parseFloat(r[1])) && parseFloat(r[1]) > 0);
    if (hasData) break;
  }
  scratch.getRange('A1:J20').clearContent();

  if (!values) return null;
  for (const row of values) {
    const v = parseFloat(row[1]);
    if (isFinite(v) && v > 0) return v;
  }
  return null;
}

function _getGoogleFinanceCurrentClose(symbol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let scratch = ss.getSheetByName(GFINANCE_SCRATCH_SHEET);
  if (!scratch) { scratch = ss.insertSheet(GFINANCE_SCRATCH_SHEET); scratch.hideSheet(); }
  scratch.getRange('A1').clearContent();
  scratch.getRange('A1').setFormula(`=GOOGLEFINANCE("${symbol}","price")`);
  SpreadsheetApp.flush();

  let val = null;
  for (let i = 0; i < 6; i++) {
    Utilities.sleep(400);
    const v = parseFloat(scratch.getRange('A1').getValue());
    if (isFinite(v) && v > 0) { val = v; break; }
  }
  scratch.getRange('A1').clearContent();
  return val;
}

function getBenchmarkComparisonData(period) {
  const startTime = Date.now();
  const TIME_BUDGET_MS = 25000; // ตัดจบก่อน client timeout (25 วิ)

  try {
    const goal = getFinancialGoalData();
    if (!goal || !goal.success) return { success: false, error: 'โหลดข้อมูล Financial Goal ไม่สำเร็จ' };

    let startDate, portfolioReturnPct, periodLabel;

    if (period === 'ytd') {
      const now = new Date();
      startDate = new Date(now.getFullYear(), 0, 1);
      periodLabel = 'YTD (ตั้งแต่ต้นปีนี้)';

      const history = getPortfolioValueHistory();
      const startPoint = (history.points || [])[0];
      if (!startPoint) {
        return { success: false, error: 'ยังไม่มีข้อมูลมูลค่าพอร์ตย้อนหลัง — กดเริ่มเก็บข้อมูลที่กราฟด้านบนก่อน' };
      }
      portfolioReturnPct = ((goal.totalAsset - startPoint.value) / startPoint.value) * 100;
      periodLabel += ` (จาก ${startPoint.date})`;
    } else {
      const now = new Date();
      startDate = new Date(now.getTime() - goal.years * 365.25 * 24 * 60 * 60 * 1000);
      periodLabel = `ตั้งแต่เริ่มลงทุน (~${goal.years.toFixed(1)} ปีที่แล้ว)`;
      portfolioReturnPct = goal.netInvest > 0 ? ((goal.totalAsset - goal.netInvest) / goal.netInvest) * 100 : null;
    }

    if (portfolioReturnPct === null) return { success: false, error: 'คำนวณผลตอบแทนพอร์ตไม่ได้' };

    const benchmarks = [];
    let partial = false;

    for (const [key, symbol] of Object.entries(BENCHMARK_SYMBOLS)) {
      const name = key === 'SET' ? 'SET Index' : 'S&P 500';

      if (Date.now() - startTime > TIME_BUDGET_MS) {
        benchmarks.push({ name, returnPct: null, beat: null, note: 'หมดเวลาดึงข้อมูล' });
        partial = true;
        continue;
      }

      try {
        const startPrice = _getGoogleFinanceCloseNear(symbol, startDate);
        const currentPrice = _getGoogleFinanceCurrentClose(symbol);
        if (startPrice && currentPrice) {
          const retPct = ((currentPrice - startPrice) / startPrice) * 100;
          benchmarks.push({ name, returnPct: retPct, beat: portfolioReturnPct > retPct });
        } else {
          benchmarks.push({ name, returnPct: null, beat: null, note: 'ดึงราคาไม่สำเร็จ (ตรวจ symbol ใน BENCHMARK_SYMBOLS)' });
        }
      } catch (symErr) {
        benchmarks.push({ name, returnPct: null, beat: null, note: 'ผิดพลาด: ' + symErr.message });
        partial = true;
      }
    }

    return { success: true, periodLabel, portfolioReturnPct, benchmarks, partial };
  } catch (e) {
    logError('getBenchmarkComparisonData', e);
    return { success: false, error: e.message };
  }
}


// ========================================
// data.gs — ดึงข้อมูลจาก Sheet จริง
// ========================================

// ----------------------------------------
// FX Rate จาก FX_Alert!B1
// ----------------------------------------
function getFxRate() {
  return Number(getCell(SHEETS.FX, FX_CELL.RATE)) || 33.4;
}

// ----------------------------------------
// Cash Balances จาก Cash Flow
// ----------------------------------------
function getCashBalances() {
  const sheet = getSheet(SHEETS.CASH);
  const BLOCK_START_ROW = 6;
  const BLOCK_START_COL = 3;
  const data = sheet.getRange(BLOCK_START_ROW, BLOCK_START_COL, 10, 4).getValues();

  const cell = (a1) => {
    const colLetter = a1.match(/[A-Z]+/)[0];
    const rowNum    = parseInt(a1.match(/\d+/)[0], 10);
    const colNum    = colLetter.charCodeAt(0) - 64;
    return data[rowNum - BLOCK_START_ROW][colNum - BLOCK_START_COL];
  };

  return {
    thb:      Number(cell(CASH_CELL.CASH_THB))   || 0,
    usd:      Number(cell(CASH_CELL.CASH_USD))   || 0,
    total:    Number(cell(CASH_CELL.CASH_TOTAL)) || 0,
    deposit:  Number(cell(CASH_CELL.DEPOSIT))    || 0,
    withdraw: Number(cell(CASH_CELL.WITHDRAW))   || 0,
    totalCost:(Number(cell(CASH_CELL.COST_THB))  || 0) +
              (Number(cell(CASH_CELL.FUND_COST)) || 0),
    exchangeTHB: Number(cell('C11')) || 0,
    totalUSDIn:  Number(cell('F6'))  || 0
  };
}

// ----------------------------------------
// Holdings US/TH
// ----------------------------------------
function getHoldings(sheetName) {
  const sheet   = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.HOLD) return [];

  const numRows = lastRow - START_ROW.HOLD + 1;
  const rows    = sheet.getRange(
    START_ROW.HOLD, 1, numRows, 17
  ).getValues();

  return rows.filter(row => {
    const ticker       = row[HOLD_COL.TICKER - 1];
    const sharesRemain = Number(row[HOLD_COL.SHARES_REMAIN - 1]) || 0;
    return ticker && sharesRemain > 0;
  }).map(row => ({
    ticker:       String(row[HOLD_COL.TICKER       - 1]).trim(),
    sharesRemain: Number(row[HOLD_COL.SHARES_REMAIN- 1]) || 0,
    avgCost:      Number(row[HOLD_COL.AVG_COST     - 1]) || 0,
    totalCost:    Number(row[HOLD_COL.TOTAL_COST   - 1]) || 0,
    priceNow:     Number(row[HOLD_COL.PRICE_NOW    - 1]) || 0,
    valueNow:     Number(row[HOLD_COL.VALUE_NOW    - 1]) || 0,
    unrealizedPL: Number(row[HOLD_COL.UNREALIZED_PL - 1]) || 0,
    unrealizedPct:Number(row[HOLD_COL.UNREALIZED_PCT- 1]) || 0,
    marketFull:   String(row[HOLD_COL.MARKET       - 1] || ""),
    dcaTarget:    Number(row[HOLD_COL.DCA_TARGET   - 1]) || 0,
    dcaBudget:    Number(row[HOLD_COL.DCA_BUDGET   - 1]) || 0
  }));
}

function getActiveHoldings(sheetName) {
  return getHoldings(sheetName);
}

// ----------------------------------------
// Fund Holdings จาก 🏛️​ 💼 Holdings
// ----------------------------------------
function getFundHoldings() {
  const sheet   = getSheet(SHEETS.FUND_HOLD);
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW.FUND) return [];

  const numRows = lastRow - START_ROW.FUND + 1;
  const rows    = sheet.getRange(
    START_ROW.FUND, 1, numRows, 12
  ).getValues();

  return rows.filter(row => {
    const name         = row[FUND_COL.NAME        - 1];
    const unitsRemain  = Number(row[FUND_COL.UNITS_REMAIN - 1]) || 0;
    return name && unitsRemain > 0;
  }).map(row => ({
    name:         String(row[FUND_COL.NAME         - 1]).trim(),
    unitsRemain:  Number(row[FUND_COL.UNITS_REMAIN - 1]) || 0,
    avgCost:      Number(row[FUND_COL.AVG_COST     - 1]) || 0,
    totalCost:    Number(row[FUND_COL.TOTAL_COST   - 1]) || 0,
    navNow:       Number(row[FUND_COL.NAV_NOW      - 1]) || 0,
    valueNow:     Number(row[FUND_COL.VALUE_NOW    - 1]) || 0,
    unrealizedPL: Number(row[FUND_COL.UNREALIZED_PL - 1]) || 0,
    unrealizedPct:Number(row[FUND_COL.UNREALIZED_PCT- 1]) || 0,
    status:       String(row[FUND_COL.STATUS       - 1] || "")
  }));
}

// ----------------------------------------
// Realized P&L จาก Realized sheet S3
// ----------------------------------------
function getRealizedPnL() {
  // ดึงจาก US sheet (ทั้ง 2 sheet มีค่า summary เหมือนกัน)
  const sheet = getSheet(SHEETS.US_REAL);
  return {
    thNet:    Number(sheet.getRange(REALIZED_CELL.TH_NET_THB).getValue())  || 0,
    usNet:    Number(sheet.getRange(REALIZED_CELL.US_NET_USD).getValue())  || 0,
    fundNet:  Number(sheet.getRange(REALIZED_CELL.FUND_NET_THB).getValue()) || 0, // ✅
    totalTHB: Number(sheet.getRange(REALIZED_CELL.TOTAL_THB).getValue())   || 0
  };
}

// ----------------------------------------
// Dividend YTD จาก บันทึกปันผล!M4
// ----------------------------------------
function getDividendYTD() {
  try {
    const sheet = getSheet(SHEETS.DIV);
    return {
      thTHB:    Number(sheet.getRange("V7").getValue())  || 0,
      usUSD:    Number(sheet.getRange("W7").getValue())  || 0,
      totalTHB: Number(sheet.getRange("X7").getValue())  || 0,
      target:   Number(sheet.getRange("AA7").getValue()) || 0  // ✅ เพิ่ม target
    };
  } catch (e) {
    logError("getDividendYTD", e);
    return { thTHB: 0, usUSD: 0, totalTHB: 0, target: 0 };
  }
}

// ----------------------------------------
// ดึงปันผลรายเดือน
// ----------------------------------------

function getDividendMonthly(year, month) {
  try {
    const sheet   = getSheet(SHEETS.DIV);
    const lastRow = sheet.getLastRow();

    // ✅ return [] ถ้าไม่มีข้อมูล
    if (!sheet || lastRow < START_ROW.DIV) return [];

    const numRows = lastRow - START_ROW.DIV + 1;
    const rows    = sheet.getRange(
      START_ROW.DIV, 1, numRows, 17
    ).getValues();

    const now = new Date();
    const yr  = year  || now.getFullYear();
    const mo  = month || (now.getMonth() + 1);

    const result = rows.filter(row => {
      const payDate = row[DIV_COL.PAY_DATE - 1];
      const ticker  = row[DIV_COL.TICKER   - 1];
      if (!ticker || !(payDate instanceof Date)) return false;
      return payDate.getFullYear() === yr &&
             (payDate.getMonth() + 1) === mo;
    }).map(row => ({
      payDate:  row[DIV_COL.PAY_DATE  - 1],
      ticker:   String(row[DIV_COL.TICKER   - 1] || ""),
      company:  String(row[DIV_COL.COMPANY  - 1] || ""),
      market:   String(row[DIV_COL.MARKET   - 1] || ""),
      perShare: Number(row[DIV_COL.PER_SHARE- 1]) || 0,
      amt:      Number(row[DIV_COL.AMT      - 1]) || 0,
      currency: String(row[DIV_COL.CURRENCY - 1] || "THB"),
      netTHB:   Number(row[DIV_COL.NET_THB  - 1]) || 0
    }));

    return result; // ✅ return array เสมอ

  } catch (e) {
    logError("getDividendMonthly", e);
    return []; // ✅ return [] ถ้า error
  }
}


// ----------------------------------------
// Master_Data
// ----------------------------------------

function getMasterData() {
  const sheet = getSheet(MASTER_DATA.SHEET);
  return {
    startDate:   sheet.getRange(MASTER_DATA.START_DATE).getValue(),
    portTarget:  Number(sheet.getRange(MASTER_DATA.PORT_TARGET).getValue())   || 0,
    initialInvest: Number(sheet.getRange(MASTER_DATA.INITIAL_INVEST).getValue()) || 0
  };
}

function getBenchmarks() {
  const sheet = getSheet(MASTER_DATA.SHEET);
  const result = [];
  for (let r = MASTER_DATA.BENCH_START_ROW; r <= 6; r++) {
    const asset  = sheet.getRange(r, MASTER_DATA.BENCH_COL_ASSET).getValue();
    const name   = sheet.getRange(r, MASTER_DATA.BENCH_COL_NAME).getValue();
    const retVal = sheet.getRange(r, MASTER_DATA.BENCH_COL_RETURN).getValue();
    if (asset && name) {
      result.push({
        asset:  String(asset),
        name:   String(name),
        return: Number(retVal) || 0
      });
    }
  }
  return result;
}






// ----------------------------------------
// collectAllData — รวมทุกอย่าง
// ----------------------------------------
function collectAllData() {
  const fxRate = getFxRate();
  const cash   = getCashBalances();

  const usHoldings   = getHoldings(SHEETS.US_HOLD);
  const thHoldings   = getHoldings(SHEETS.TH_HOLD);
  const fundHoldings = getFundHoldings();    // ✅ รวมกองทุน

  const usPortUSD    = usHoldings.reduce((s, h) => s + h.valueNow, 0);
  const usPortTHB    = usPortUSD * fxRate;
  const thPortTHB    = thHoldings.reduce((s, h) => s + h.valueNow, 0);
  const fundPortTHB  = fundHoldings.reduce((s, h) => s + h.valueNow, 0); // ✅

  const usUnrealUSD  = usHoldings.reduce((s, h) => s + h.unrealizedPL, 0);
  const usUnrealTHB  = usUnrealUSD * fxRate;
  const thUnrealTHB  = thHoldings.reduce((s, h) => s + h.unrealizedPL, 0);
  const fundUnrealTHB= fundHoldings.reduce((s, h) => s + h.unrealizedPL, 0); // ✅


  const totalPortTHB   = thPortTHB + usPortTHB + fundPortTHB; // ✅ รวมกองทุน
  const totalUnrealTHB = thUnrealTHB + usUnrealTHB + fundUnrealTHB; // ✅

  const realized = getRealizedPnL();
  const div      = getDividendYTD();

  const usSorted = [...usHoldings].sort((a, b) => b.unrealizedPct - a.unrealizedPct);
  const thSorted = [...thHoldings].sort((a, b) => b.unrealizedPct - a.unrealizedPct);

  return {
    fxRate,
    // Cash
    cashTHB:    cash.thb,
    cashUSD:    cash.usd,
    cashTotal:  cash.total,
    deposit:    cash.deposit,
    withdraw:   cash.withdraw,
    totalCost:  cash.totalCost,
    netInvest:  cash.deposit - cash.withdraw,
    exchangeTHB: cash.exchangeTHB,
    totalUSDIn:  cash.totalUSDIn,

    // Holdings
    usHoldings, thHoldings, fundHoldings,
    // Portfolio values
    usPortUSD, usPortTHB,
    thPortTHB,
    fundPortTHB,
    totalPortTHB,
    // Unrealized
    usUnrealUSD, usUnrealTHB,
    thUnrealTHB,
    fundUnrealTHB,
    totalUnrealTHB,
    // Realized
    thRealTHB:   realized.thNet,
    usRealUSD:   realized.usNet,
    usRealTHB:   realized.usNet * fxRate,
    fundRealTHB: realized.fundNet,   // ✅ เพิ่มบรรทัดนี้
    totalRealTHB:realized.totalTHB,
    
    // Dividend
    divTHB:      div.thTHB,
    divUSD:      div.usUSD,
    divTotalTHB: div.totalTHB,
    // Sorted
    usTop3:  usSorted.slice(0, 3),
    usWorst: usSorted[usSorted.length - 1] || null,
    thTop3:  thSorted.slice(0, 3),
    thWorst: thSorted[thSorted.length - 1] || null
  };
}


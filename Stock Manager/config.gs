
// ========================================
// config.gs — อัปเดตให้ตรงกับ sheet จริง
// ========================================

const SHEETS = {
  US_TRANS:   "🇺🇸 หุ้นสหรัฐ (USD)",
  TH_TRANS:   "🇹🇭 หุ้นไทย (THB)",
  US_HOLD:    "💼 Holdings",
  TH_HOLD:    "🇹🇭  💼 Holdings",
  FUND_HOLD:  "🏛️​ 💼 Holdings",  // ✅ มี zero-width space
  US_REAL:    "📈 Realized P&L",
  TH_REAL:    "🇹🇭 📈 Realized P&L",
  DIV:        "📝 บันทึกปันผล",
  CASH:       "💵 Cash Flow",
  FX:         "FX_Alert",
  BACKEND_ATR:"Backend_ATR",
  ATR_US:     "📊 ATR_Portfolio US",
  ATR_TH:     "📊 ATR_Portfolio TH"
  
};

const START_ROW = {
  HOLD:     7,   // Holdings data เริ่ม row 7
  FUND:     4,   // Fund Holdings data เริ่ม row 4
  DIV:      7,   // Dividend data เริ่ม row 7
  REALIZED: 7,   // Realized P&L data เริ่ม row 7
  ATR_US:   6,
  ATR_TH:   6
};




// Holdings US/TH col (1-based)
const HOLD_COL = {
  TICKER:         2,   // B
  SHARES_BUY:     3,   // C
  SHARES_SOLD:    4,   // D
  SHARES_REMAIN:  5,   // E
  AVG_COST:       6,   // F
  TOTAL_COST:     7,   // G
  PRICE_NOW:      8,   // H ← ราคาปัจจุบัน (GAS เขียน)
  VALUE_NOW:      9,   // I
  UNREALIZED_PL:  10,  // J
  UNREALIZED_PCT: 11,  // K
  MARKET:         12,  // L (format "NYSEARCA:VOO")
  TODAY_HIGH:     15,  // O (TH only)
  DCA_TARGET:     16,  // P
  DCA_BUDGET:     17   // Q
};

// Fund Holdings col (1-based)
const FUND_COL = {
  NAME:          2,   // B
  UNITS_BUY:     3,   // C
  UNITS_SOLD:    4,   // D
  UNITS_REMAIN:  5,   // E
  AVG_COST:      6,   // F
  TOTAL_COST:    7,   // G
  NAV_NOW:       8,   // H ← GAS เขียน NAV ที่นี่
  VALUE_NOW:     9,   // I
  UNREALIZED_PL: 10,  // J
  UNREALIZED_PCT:11,  // K
  STATUS:        12   // L
};

// Cash Flow cells
const CASH_CELL = {
  DEPOSIT:     "C6",
  WITHDRAW:    "C7",
  COST_THB:    "C8",
  COST_USD:    "F8",
  FUND_COST:   "C12",
  CASH_THB:    "C14",   // ✅ เงินสดคงเหลือ THB
  CASH_USD:    "F14",   // ✅ เงินสดคงเหลือ USD
  CASH_TOTAL:  "F15"    // ✅ เงินสดคงเหลือรวม THB
};

// FX Alert cells
const FX_CELL = {
  RATE:        "B1",
  BUY1:        "B2",
  BUY2:        "B3",
  SELL:        "B4",
  FLAG_BUY1:   "B6",
  FLAG_BUY2:   "B7",
  FLAG_SELL:   "B8",
  LAST_CHECK:  "B10"
};

// Realized P&L summary cells (ทั้ง US และ TH sheet มี cell เดียวกัน)
const REALIZED_CELL = {
  TH_NET_THB:  "S2",   // P&L ไทย สุทธิ (THB) = 12.88
  US_NET_USD:  "U2",   // P&L สหรัฐ สุทธิ (USD) = -17.05
  TOTAL_THB:   "S3",   // P&L รวมทั้งหมด (THB) = -556.70
  FX_RATE:     "P2",    // อัตราแลกเปลี่ยนที่ใช้คำนวณ
  FUND_NET_THB:"W2"    // P&L กองทุน สุทธิ (USD)
};

// Dividend summary cells
const DIV_CELL = {
  // ปีปัจจุบัน (2026) row 7
  TH_THB:    "V7",   // ปันผลไทย (THB) = 0.00
  US_USD:    "W7",   // ปันผลสหรัฐ (USD) = 97.63
  TOTAL_THB: "X7",   // รวม (THB) = 3,258.51 ✅
  TARGET:    "AA7",  // เป้าหมาย = 2,000

  // รวมทุกปี row 12
  ALL_TH_THB:    "V12",  // รวมปันผลไทย
  ALL_US_USD:    "W12",  // รวมปันผลสหรัฐ
  ALL_TOTAL_THB: "X12"   // รวมทั้งหมด THB
};


// Dividend log columns (1-based)
const DIV_COL = {
  NO:        2,   // B
  XD_DATE:   3,   // C
  PAY_DATE:  4,   // D
  TICKER:    5,   // E
  COMPANY:   6,   // F
  MARKET:    7,   // G
  ROUND:     8,   // H
  YEAR:      9,   // I
  PER_SHARE: 10,  // J
  AMT:       11,  // K ← ปันผลรับรวม
  SHARES:    12,  // L
  CURRENCY:  13,  // M
  TAX_USD:   14,  // N
  TAX_PCT:   15,  // O
  NET_THB:   16   // P ← ปันผลสุทธิ (THB)
};

// ✅ เพิ่มใน config.gs
const ATR_SHEETS = {
  US: "📊 ATR_Portfolio US",
  TH: "📊 ATR_Portfolio TH"
};

const ATR_START_ROW = {
  US: 6,
  TH: 6
};

const FUND_HOLD_COL = {
  NAME:          2,   // B
  UNITS_BUY:     3,   // C
  UNITS_SOLD:    4,   // D
  UNITS_REMAIN:  5,   // E
  AVG_COST:      6,   // F
  TOTAL_COST:    7,   // G
  NAV_NOW:       8,   // H
  VALUE_NOW:     9,   // I
  UNREALIZED_PL: 10,  // J
  UNREALIZED_PCT:11,  // K
  STATUS:        12   // L
  
};

// ========================================
// config.gs — เพิ่ม ATR_COL
// ========================================

const ATR_COL = {
  SYMBOL:         0,   // A — ชื่อหุ้น
  TRADE_STYLE:    1,   // B — Trade Style
  BUY_PRICE:      2,   // C — Buy Price
  HIGHEST_CLOSE:  3,   // D — Highest Close
  ATR:            4,   // E — ATR (14)
  MULTIPLIER:     5,   // F — Multiplier
                       // G = ว่าง (index 6)
  TRAILING_STOP:  7,   // H — Trailing Stop
  STATUS:         8,   // I — Status
  MIN_PROFIT:     9,   // J — ค่าไซขั้นต่ำที่จุด stop
  SHARES:         10,  // K — จำนวนหุ้น
  PRICE_NOW:      11,  // L — ราคาปัจจุบัน
  PL_NOW:         12,  // M — กำไร/ขาดทุนปัจจุบัน
  TAKE_PROFIT:    13,  // N — Take Profit
  RR:             14,  // O — ค่า R/R
  PROFIT_TARGET:  15,  // P — กำไรเป้า
  TOTAL_PROFIT:   16,  // Q — กำไรทั้งหมด
  PCT_PROFIT:     17,  // R — %กำไรเป้า
  MAX_RISK:       18,  // S — ขาดทุนสูงสุด (Max Risk)
  TOTAL_LOSS:     19,  // T — ขาดทุนทั้งหมด
  PCT_LOSS:       20,  // U — %ขาดทุนสูงสุด
  RR_REAL:        21,  // V — R/R จริง
  SUMMARY_STATUS: 22,  // W — สถานะ ("ขายที่" / "ถือต่อ")
  STOP_DISTANCE:  23,  // X — Stop Distance %
  RISK_STATUS:    24   // Y — Risk Status
};


const MASTER_DATA = {
  SHEET:          "Master_Data",
  START_DATE:     "B2",
  PORT_TARGET:    "B3",
  INITIAL_INVEST: "B4",
  // Benchmark rows 3-6, col D=สินทรัพย์, E=Benchmark, F=ผลตอบแทน YTD
  BENCH_START_ROW: 3,
  BENCH_COL_ASSET: 4,   // D
  BENCH_COL_NAME:  5,   // E
  BENCH_COL_RETURN:6    // F
};


// Backend_ATR (1-based col)
const BACKEND_ATR = {
  SHEET:     "Backend_ATR",
  US_TICKER: 1,   // A
  US_ATR:    2,   // B
  TH_TICKER: 5,   // E
  TH_ATR:    6,   // F
  START_ROW: 2
};








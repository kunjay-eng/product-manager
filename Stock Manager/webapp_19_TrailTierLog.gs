// ══════════════════════════════════════════════════════════
// webapp_19_TrailTierLog.gs —
// ------------------------------------------------------------
// ══════════════════════════════════════════════════════════


const TRAIL_LOG_SHEET = 'TrailTierLog';

function _findTrailLogRow(sheet, ticker, market, tierKey) {
  if (sheet.getLastRow() < 2) return -1;
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // Ticker,Market,TierKey
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === ticker && data[i][1] === market && data[i][2] === tierKey) {
      return i + 2;
    }
  }
  return -1;
}

function _getExecutedTiers(ticker, market, currentHighestHigh, resetPct) {
  const sheet = getSheet(TRAIL_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return new Map();

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const executed = new Map();

  data.forEach(row => {
    const [logId, t, m, tierKey, execDate, highAtExec, stopPrice, triggerPrice, sharesSold, avgCost, atr, multiplier, currentPrice] = row;
    if (String(t).toUpperCase() !== ticker || m !== market) return;
    // High ใหม่สูงกว่าที่บันทึกไว้เกิน resetPct% → รอบใหม่ ไม่นับว่า executed แล้ว (auto-reset)
    if (currentHighestHigh > Number(highAtExec) * (1 + resetPct / 100)) return;
    executed.set(tierKey, { logId, execDate, highAtExec, stopPrice, triggerPrice, sharesSold, avgCost, atr, multiplier, currentPrice });
  });

  return executed;
}

/**
 * รับแค่ ticker/market/tierKey จาก UI — Backend คำนวณค่าจริงเองทั้งหมดจาก
 * getPortfolioTrailingStop() ล่าสุด ไม่เชื่อค่าที่ UI ส่งมา (Single Source of Truth)
 * Upsert: มี tier เดิม → update ทับ, ยังไม่มี → insert ใหม่ (ใช้ LogId เดิมถ้ามี)
 */
function markTrailTierExecuted(ticker, market, tierKey) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();

    const fresh = getPortfolioTrailingStop(ticker, market);
    if (!fresh.success) return { success: false, error: 'คำนวณข้อมูลล่าสุดไม่ได้: ' + fresh.error };

    const tier = fresh.tiers.find(t => t.key === tierKey);
    if (!tier) return { success: false, error: 'ไม่พบ tier: ' + tierKey };

    const sheet = getSheet(TRAIL_LOG_SHEET);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + TRAIL_LOG_SHEET);

    const existingRow = _findTrailLogRow(sheet, ticker, market, tierKey);
    const logId = existingRow > 0
      ? sheet.getRange(existingRow, 1).getValue()
      : Utilities.getUuid();

    const rowValues = [
      logId, ticker, market, tierKey, new Date(),
      fresh.highestHighSinceEntry, tier.stopPrice, fresh.triggerPrice, tier.sharesToSell,
      fresh.avgCost, fresh.atr, tier.multiplier, fresh.currentPrice
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, 13).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
    return { success: true, logId, snapshot: rowValues };
  } catch (e) {
    logError('markTrailTierExecuted', e);
    return { success: false, error: e.message };
  }
}

/** รีเซ็ตเฉพาะ tier ที่ระบุ ไม่แตะ tier อื่น */
function resetTrailTierState(ticker, market, tierKey) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const sheet = getSheet(TRAIL_LOG_SHEET);
    if (!sheet) throw new Error('ไม่พบ sheet: ' + TRAIL_LOG_SHEET);

    const row = _findTrailLogRow(sheet, ticker, market, tierKey);
    if (row > 0) sheet.deleteRow(row);

    return { success: true };
  } catch (e) {
    logError('resetTrailTierState', e);
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════
// webapp_21_flexible_buy_plan.gs — แผนแบ่งไม้แบบยืดหยุ่น
// ไม้ 1 อัตโนมัติ (ราคาจริง/ราคาปัจจุบัน) — ไม้ 2,3 ผู้ใช้กำหนด %+ราคาเอง
// ══════════════════════════════════════════════════════════

const BUY_PLAN_REF_PRICE_COL = 15; // O — ราคาอ้างอิงตอนสร้างแผน (ใช้แทนราคาซื้อจริงถ้ายังไม่เคยถือ)

// ── หาราคาอ้างอิงไม้ 1 (ราคาซื้อจริงถ้าถืออยู่แล้ว / ราคาปัจจุบันถ้ายังไม่ถือ)
//    แยกออกมาให้ modal preview เรียกดูก่อนบันทึกได้ ไม่ต้องเดา ──
function _resolveBuyPlanReferencePrice(ticker, market) {
  const logSheetName = (market === 'TH') ? SHEETS.TH_TRANS : SHEETS.US_TRANS;
  const logRows = _getSheetDataCached(logSheetName);
  const firstBuy = logRows
    .filter(r => String(r[2] || '').trim().toUpperCase().replace(/_C\d+$/i, '') === ticker && r[3] === 'ซื้อ')
    .map(r => ({ date: r[1], price: parseFloat(r[5]) }))
    .filter(x => x.date && !isNaN(x.price))
    .sort((a, b) => new Date(a.date) - new Date(b.date))[0];

  if (firstBuy) return { price: firstBuy.price, source: 'first_buy' };

  const quote = _wlFetchYahooQuote(ticker, market);
  if (!quote) return { price: null, source: null };
  return { price: quote.price, source: 'current_price' };
}

// ── Preview ราคาอ้างอิงก่อนบันทึก — ใช้ตอนเปิด modal ──
function getBuyPlanReferencePriceHint(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    const ref = _resolveBuyPlanReferencePrice(ticker, market);
    if (ref.price === null) return { success: false, error: 'ไม่พบราคาของ ' + ticker };
    return { success: true, referencePrice: ref.price, source: ref.source };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveFlexibleBuyPlan(ticker, market, totalBudget, legsInput, note, direction) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = String(market || '').trim().toUpperCase();
    direction = (direction === 'up') ? 'up' : 'down'; // default = down (DCA ย่อ แบบเดิม)
    totalBudget = parseFloat(totalBudget);
    if (!totalBudget || totalBudget <= 0) return { success: false, error: 'กรุณาระบุงบลงทุนรวมที่ถูกต้อง' };
    if (!legsInput || !legsInput.length) return { success: false, error: 'กรุณากำหนดไม้ 2 อย่างน้อย 1 ไม้' };
    if (legsInput.length > 2) return { success: false, error: 'รองรับสูงสุด 3 ไม้ (ไม้ 1 อัตโนมัติ + กำหนดเองอีก 2 ไม้)' };

    const extraPctSum = legsInput.reduce((s, l) => s + (parseFloat(l.pct) || 0), 0);
    if (extraPctSum >= 100) return { success: false, error: '% รวมของไม้ 2-3 ต้องน้อยกว่า 100% (ต้องเหลือให้ไม้ 1)' };
    if (extraPctSum <= 0) return { success: false, error: 'กรุณาระบุ % ของไม้ 2-3 ให้มากกว่า 0' };
    for (const l of legsInput) {
      if (!l.price || parseFloat(l.price) <= 0) return { success: false, error: 'กรุณาระบุราคาเป้าหมายของทุกไม้ให้ครบ' };
    }

    const ref = _resolveBuyPlanReferencePrice(ticker, market);
    if (ref.price === null) return { success: false, error: 'ไม่พบราคาปัจจุบันของ ' + ticker };
    const referencePrice = ref.price;

    // ── VALIDATION: ทิศทางเดียวกันตลอดทั้งแผน ตามที่ผู้ใช้เลือก
    //    direction='down' → แต่ละไม้ต้องถูกกว่าไม้ก่อนหน้า (DCA ย่อ)
    //    direction='up'   → แต่ละไม้ต้องแพงกว่าไม้ก่อนหน้า (Pyramid ยืนยันขึ้น) ──
    let prevPrice = referencePrice;
    let prevLabel = 'ไม้ 1 (ราคาอ้างอิง ' + fmtNumServer(referencePrice) + ')';
    for (let i = 0; i < legsInput.length; i++) {
      const p = parseFloat(legsInput[i].price);
      const legLabel = 'ไม้ ' + (i + 2);
      const invalid = direction === 'down' ? (p >= prevPrice) : (p <= prevPrice);
      if (invalid) {
        const requirement = direction === 'down' ? 'ต่ำกว่า' : 'สูงกว่า';
        return {
          success: false,
          error: `ราคาเป้าหมาย${legLabel} (${fmtNumServer(p)}) ต้อง${requirement}${prevLabel} — แผนนี้ตั้งเป็น "${direction === 'down' ? 'ไล่ซื้อตอนราคาย่อ' : 'ไล่ซื้อตอนราคายืนยันขึ้น'}" แต่ละไม้ต้องเรียงทิศทางเดียวกันตลอด`
        };
      }
      prevPrice = p;
      prevLabel = legLabel + ' (' + fmtNumServer(p) + ')';
    }

    const leg1Pct = 100 - extraPctSum;
    const legsForSheet = [{ pct: leg1Pct, triggerPct: 0 }];
    legsInput.forEach(l => {
      const price = parseFloat(l.price);
      // triggerPct เก็บเป็นค่าสัมบูรณ์เสมอ (ระยะห่าง % จากไม้ 1) — ทิศทางดูจาก field 'direction'
      // แยกต่างหาก ไม่ผสมกันใน trigger% เดียว กัน logic เดิมที่ใช้ trigger% แบบ "ย่อลงเสมอ" พัง
      const triggerPct = Math.abs((referencePrice - price) / referencePrice) * 100;
      legsForSheet.push({ pct: parseFloat(l.pct), triggerPct: _taxRound(triggerPct, 2) });
    });

    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= BUY_PLAN_SHEET.START_ROW) {
      const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
      const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 2).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toUpperCase() === ticker && String(rows[i][1]).trim().toUpperCase() === market) {
          targetRow = BUY_PLAN_SHEET.START_ROW + i; break;
        }
      }
    }
    if (targetRow === -1) targetRow = _getNextEmptyRow(sheet, 1, BUY_PLAN_SHEET.START_ROW);

    sheet.getRange(targetRow, 1).setValue(ticker);
    sheet.getRange(targetRow, 2).setValue(market);
    sheet.getRange(targetRow, 3).setValue('price');
    sheet.getRange(targetRow, 4).setValue(totalBudget);
    legsForSheet.forEach((leg, i) => {
      sheet.getRange(targetRow, 5 + i * 2).setValue(leg.pct);
      sheet.getRange(targetRow, 6 + i * 2).setValue(leg.triggerPct);
    });
    for (let i = legsForSheet.length; i < 3; i++) {
      sheet.getRange(targetRow, 5 + i * 2).setValue('');
      sheet.getRange(targetRow, 6 + i * 2).setValue('');
    }
    sheet.getRange(targetRow, 13).setValue(new Date());
    sheet.getRange(targetRow, 14).setValue(note || 'แผนแบ่งไม้แบบกำหนดเอง');
    sheet.getRange(targetRow, BUY_PLAN_REF_PRICE_COL).setValue(referencePrice);
    sheet.getRange(targetRow, 16).setValue(direction); // ← คอลัมน์ P ใหม่ เก็บทิศทางแผน

    _logBuyPlanHistory(ticker, market, 'created', { totalBudget, leg1Pct, legs: legsForSheet, referencePrice, direction });

    return { success: true, row: targetRow, referencePrice, leg1Pct, direction };
  } catch (e) {
    logError('saveFlexibleBuyPlan', e);
    return { success: false, error: e.message };
  }
}


// ── ฟอร์แมตตัวเลขง่ายๆ ใช้ในข้อความ error ฝั่ง server (ไม่ต้องพึ่ง fmtNum ฝั่ง client) ──
function fmtNumServer(v) {
  return Number(v).toFixed(2);
}


/** แก้ไขงบลงทุนรวม — ไม่แตะ %/ราคาเป้าหมายเดิม */
function updateBuyPlanTarget(ticker, market, newBudget) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = String(market || '').trim().toUpperCase();
    newBudget = parseFloat(newBudget);
    if (!newBudget || newBudget <= 0) return { success: false, error: 'กรุณาระบุงบลงทุนที่ถูกต้อง' };

    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < BUY_PLAN_SHEET.START_ROW) return { success: false, error: 'ไม่พบแผนของ ' + ticker };

    const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
    const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() === ticker && String(rows[i][1]).trim().toUpperCase() === market) {
        sheet.getRange(BUY_PLAN_SHEET.START_ROW + i, 4).setValue(newBudget);
         _logBuyPlanHistory(ticker, market, 'target_edited', { newBudget });

        return { success: true, newBudget };
      }
    }
    return { success: false, error: 'ไม่พบแผนของ ' + ticker + ' — สร้างแผนใหม่ก่อน' };
  } catch (e) {
    logError('updateBuyPlanTarget', e);
    return { success: false, error: e.message };
  }
}

/** เทียบงบที่เหลือต้องลงทุนตามแผน กับเงินสดจริงที่มี (แยก THB/USD) */
function getBuyPlanCashStatus(ticker, market) {
  try {
    const planResult = getBuyPlanForTicker(ticker, market);
    if (!planResult.success) return { success: false, error: planResult.error };
    if (planResult.planType !== 'price') return { success: false, error: 'ใช้ได้เฉพาะแผนแบบ Price-based' };

    const steps = planResult.steps || [];
    const totalBudget = planResult.budget || 0;
    const spent = steps.filter(s => s.status === 'done').reduce((sum, s) => sum + (s.actualBudgetSpent || 0), 0);
    const remainingBudget = Math.max(0, totalBudget - spent);

    const cash = getCashBalances();
    const cur = market === 'TH' ? 'THB' : 'USD';
    const cashAvailable = market === 'TH' ? cash.thb : cash.usd;
    const cashSufficient = cashAvailable >= remainingBudget;

    return {
      success: true, ticker, market, currency: cur,
      totalBudget, spent: _taxRound(spent, 2), remainingBudget: _taxRound(remainingBudget, 2),
      cashAvailable: _taxRound(cashAvailable, 2),
      cashSufficient, shortfall: _taxRound(cashSufficient ? 0 : (remainingBudget - cashAvailable), 2)
    };
  } catch (e) {
    logError('getBuyPlanCashStatus', e);
    return { success: false, error: e.message };
  }
}


// ── สแกนทุกแผนที่มีในชีต BuyPlan หาไม้ที่ canBuyNow แล้ว — ใช้แสดงในหน้า Summary
//    จำกัดเฉพาะ planType='price' เท่านั้น (DCA แบบ time-based ไม่มี concept "ราคาที่ควรซื้อ") ──
function getBuyPlanAlertsSummary() {
  try {
    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < BUY_PLAN_SHEET.START_ROW) return { success: true, alerts: [] };

    const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
    const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 3).getValues();

    const alerts = [];
    rows.forEach(r => {
      const ticker = String(r[0] || '').trim().toUpperCase();
      const market = String(r[1] || '').trim().toUpperCase();
      const planType = String(r[2] || '').trim().toLowerCase();
      if (!ticker || planType !== 'price') return;

      try {
        const plan = getBuyPlanForTicker(ticker, market);
        if (!plan.success || plan.planType !== 'price') return;

        const readyStep = (plan.steps || []).find(s => s.status !== 'done' && s.canBuyNow);
        if (readyStep) {
          alerts.push({
            ticker, market,
            legNumber: readyStep.legNumber,
            targetPrice: readyStep.targetPrice,
            currentPrice: plan.currentPrice,
            pct: readyStep.pct,
            cur: plan.cur
          });
        }
      } catch (innerErr) {
        logError('getBuyPlanAlertsSummary:' + ticker, innerErr);
        // ข้ามตัวที่ error ไป ไม่ให้ทั้งฟังก์ชันพัง
      }
    });

    return { success: true, alerts };
  } catch (e) {
    logError('getBuyPlanAlertsSummary', e);
    return { success: false, error: e.message, alerts: [] };
  }
}

// ── เช็คว่าถ้าซื้อครบทุกไม้ตามแผน จะถือกี่ % ของพอร์ตรวม เทียบ concentrationWarnPct
//    reuse getRiskSettings()/getFinancialGoalData()/getFxRate()/getHoldingsData() เดิมทั้งหมด ──
function getBuyPlanConcentrationCheck(ticker, market) {
  try {
    const plan = getBuyPlanForTicker(ticker, market);
    if (!plan.success || plan.planType !== 'price') return { success: false, error: 'ใช้ได้เฉพาะแผนแบบ Price-based' };

    const risk = getRiskSettings();
    const goal = getFinancialGoalData();
    if (!goal.success) return { success: false, error: 'ดึงมูลค่าพอร์ตไม่สำเร็จ' };

    const totalAssetTHB = goal.totalAsset;
    const fx = market === 'US' ? getFxRate() : 1;

    // ── มูลค่าที่ถืออยู่ตอนนี้ของ ticker นี้ (ถ้ามี) ──
    const holdings = getHoldingsData();
    const arr = market === 'TH' ? holdings.th : holdings.us;
    const currentHolding = (arr || []).find(h => h.ticker === ticker);
    const currentValueTHB = currentHolding ? (parseFloat(currentHolding.valueNow) || 0) * fx : 0;

    // ── งบที่ยังไม่ได้ใช้ (ไม้ที่ยัง not done) แปลงเป็น THB ──
    const remainingBudget = (plan.steps || [])
      .filter(s => s.status !== 'done')
      .reduce((sum, s) => sum + (s.budgetForLeg || 0), 0);
    const remainingBudgetTHB = remainingBudget * fx;

    const projectedValueTHB = currentValueTHB + remainingBudgetTHB;
    const currentPct = totalAssetTHB > 0 ? (currentValueTHB / totalAssetTHB) * 100 : 0;
    const projectedPct = totalAssetTHB > 0 ? (projectedValueTHB / totalAssetTHB) * 100 : 0;
    const limitPct = risk.concentrationWarnPct * 100;

    return {
      success: true,
      currentPct: _taxRound(currentPct, 1),
      projectedPct: _taxRound(projectedPct, 1),
      limitPct: _taxRound(limitPct, 1),
      overLimit: projectedPct > limitPct,
      currentValueTHB: _taxRound(currentValueTHB, 0),
      projectedValueTHB: _taxRound(projectedValueTHB, 0)
    };
  } catch (e) {
    logError('getBuyPlanConcentrationCheck', e);
    return { success: false, error: e.message };
  }
}

// ── ยกเลิกไม้ 2 หรือ 3 ที่ยังไม่ซื้อ — ลบออกจากแผนแล้วเลื่อนไม้ที่เหลือขึ้นมาแทนที่
//    ห้ามยกเลิกไม้ 1 เพราะเป็นราคาอ้างอิงคำนวณ trigger% ของไม้อื่นทั้งหมด ──
function cancelBuyPlanLeg(ticker, market, legNumber) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = String(market || '').trim().toUpperCase();
    legNumber = parseInt(legNumber, 10);
    if (legNumber === 1) return { success: false, error: 'ยกเลิกไม้ 1 ไม่ได้ เพราะเป็นราคาอ้างอิงของทั้งแผน' };

    const cfg = _readBuyPlanConfig(ticker, market);
    if (!cfg) return { success: false, error: 'ไม่พบแผนของ ' + ticker };
    if (legNumber < 1 || legNumber > cfg.legs.length) return { success: false, error: 'ไม่พบไม้ที่ ' + legNumber + ' ในแผนนี้' };

    // ── กันยกเลิกไม้ที่ซื้อไปแล้ว — เช็คสถานะจริงจาก getBuyPlanForTicker ก่อน ──
    const statusCheck = getBuyPlanForTicker(ticker, market);
    if (statusCheck.success) {
      const targetStep = statusCheck.steps.find(s => s.legNumber === legNumber);
      if (targetStep && targetStep.status === 'done') {
        return { success: false, error: 'ไม้ ' + legNumber + ' ซื้อไปแล้ว ยกเลิกไม่ได้' };
      }
    }

    // ── ลบไม้นั้นออก แล้วเลื่อนไม้ถัดไปขึ้นมาแทนที่ (index 0-based: legNumber 2 = index 1) ──
    const remainingLegs = cfg.legs.filter((_, idx) => idx !== (legNumber - 1));

    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
    const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() === ticker && String(rows[i][1]).trim().toUpperCase() === market) {
        targetRow = BUY_PLAN_SHEET.START_ROW + i; break;
      }
    }
    if (targetRow === -1) return { success: false, error: 'ไม่พบแถวแผนในชีต' };

    // ── เขียนไม้ที่เหลือกลับลงคอลัมน์ E-J เรียงใหม่ ที่เหลือเว้นว่าง ──
    for (let i = 0; i < 3; i++) {
      if (i < remainingLegs.length) {
        sheet.getRange(targetRow, 5 + i * 2).setValue(remainingLegs[i].pct);
        sheet.getRange(targetRow, 6 + i * 2).setValue(remainingLegs[i].triggerPct);
      } else {
        sheet.getRange(targetRow, 5 + i * 2).setValue('');
        sheet.getRange(targetRow, 6 + i * 2).setValue('');
      }
    }
    _logBuyPlanHistory(ticker, market, 'leg_cancelled', { legNumber, remainingLegsCount: remainingLegs.length });
    return { success: true, remainingLegsCount: remainingLegs.length };
  } catch (e) {
    logError('cancelBuyPlanLeg', e);
    return { success: false, error: e.message };
  }
}

// ── เช็คว่าแผนนี้ตั้งมานานแค่ไหนแล้ว นับจากวันที่เริ่มแผน (col M) ถึงวันนี้
//    ใช้เตือนถ้าไม่มีความเคลื่อนไหว (ไม่มีไม้ไหน done) นานเกินเกณฑ์ ──
const BUY_PLAN_STALE_DAYS_WARN = 90;  // เตือนสีเหลืองถ้าเกินนี้
const BUY_PLAN_STALE_DAYS_ALERT = 180; // เตือนสีแดงถ้าเกินนี้

function getBuyPlanAgeCheck(ticker, market) {
  try {
    const cfg = _readBuyPlanConfig(ticker, market);
    if (!cfg || !cfg.startDate) return { success: false, error: 'ไม่พบวันที่เริ่มแผน' };

    const plan = getBuyPlanForTicker(ticker, market);
    if (!plan.success) return { success: false, error: plan.error };

    const daysSinceStart = Math.floor((new Date() - cfg.startDate) / (1000 * 60 * 60 * 24));
    const hasAnyExecuted = (plan.steps || []).some(s => s.status === 'done');
    // ── เตือนเฉพาะกรณี "ตั้งแผนไว้นานแล้วแต่ยังไม่มีไม้ไหนซื้อเลย" — ถ้าซื้อไปแล้วบางไม้
    //    แปลว่าแผนยัง active อยู่จริง ไม่ใช่แผนที่ถูกลืม ไม่ต้องเตือน ──
    const isStale = !hasAnyExecuted && daysSinceStart >= BUY_PLAN_STALE_DAYS_WARN;

    let staleLevel = null;
    if (isStale) staleLevel = daysSinceStart >= BUY_PLAN_STALE_DAYS_ALERT ? 'alert' : 'warn';

    return {
      success: true,
      daysSinceStart,
      hasAnyExecuted,
      isStale,
      staleLevel,
      startDate: Utilities.formatDate(cfg.startDate, 'Asia/Bangkok', 'dd/MM/yyyy')
    };
  } catch (e) {
    logError('getBuyPlanAgeCheck', e);
    return { success: false, error: e.message };
  }
}

// ── สแกนทุกแผนหาตัวที่ stale — ใช้แสดงรวมในหน้า Summary คู่กับ getBuyPlanAlertsSummary() เดิม ──
function getStaleBuyPlansSummary() {
  try {
    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < BUY_PLAN_SHEET.START_ROW) return { success: true, stalePlans: [] };

    const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
    const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 3).getValues();

    const stalePlans = [];
    rows.forEach(r => {
      const ticker = String(r[0] || '').trim().toUpperCase();
      const market = String(r[1] || '').trim().toUpperCase();
      const planType = String(r[2] || '').trim().toLowerCase();
      if (!ticker || planType !== 'price') return;

      try {
        const ageCheck = getBuyPlanAgeCheck(ticker, market);
        if (ageCheck.success && ageCheck.isStale) {
          stalePlans.push({ ticker, market, daysSinceStart: ageCheck.daysSinceStart, staleLevel: ageCheck.staleLevel });
        }
      } catch (innerErr) {
        logError('getStaleBuyPlansSummary:' + ticker, innerErr);
      }
    });

    return { success: true, stalePlans };
  } catch (e) {
    logError('getStaleBuyPlansSummary', e);
    return { success: false, error: e.message, stalePlans: [] };
  }
}


// ══════════════════════════════════════════════════════════
// ข้อ 6: ประวัติแผนเก่า — log แยกจากชีต BuyPlan (ที่เก็บแค่แผนปัจจุบัน)
// บันทึกทุก action: created / target_edited / leg_cancelled
// ══════════════════════════════════════════════════════════
const BUY_PLAN_HISTORY_SHEET = 'BuyPlan_History';
// คอลัมน์: A=Timestamp B=Ticker C=Market D=Action E=Detail(JSON string)

function _logBuyPlanHistory(ticker, market, action, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(BUY_PLAN_HISTORY_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(BUY_PLAN_HISTORY_SHEET);
      sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Ticker', 'Market', 'Action', 'Detail']])
        .setFontWeight('bold').setBackground('#e8eaf6');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), ticker, market, action, JSON.stringify(detail || {})]);
  } catch (e) {
    logError('_logBuyPlanHistory', e); // ไม่ throw — ล็อกพังไม่ควรทำให้ action หลักพังตาม
  }
}

// ── ดึงประวัติของ ticker ตัวเดียว เรียงล่าสุดก่อน ──
function getBuyPlanHistory(ticker, market) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = String(market || '').trim().toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(BUY_PLAN_HISTORY_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, history: [] };

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    const actionLabels = { created: '📝 สร้างแผนใหม่', target_edited: '✏️ แก้ไขงบเป้าหมาย', leg_cancelled: '✕ ยกเลิกไม้',leg_edited: '✏️ แก้ไขไม้' };

    const history = rows
      .filter(r => String(r[1]).trim().toUpperCase() === ticker && String(r[2]).trim().toUpperCase() === market)
      .map(r => {
        let detail = {};
        try { detail = JSON.parse(r[4] || '{}'); } catch (e) {}
        return {
          timestamp: Utilities.formatDate(new Date(r[0]), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'),
          action: r[3],
          actionLabel: actionLabels[r[3]] || r[3],
          detail
        };
      })
      .sort((a, b) => new Date(b.timestamp.split(' ')[0].split('/').reverse().join('-') + ' ' + b.timestamp.split(' ')[1]) -
                       new Date(a.timestamp.split(' ')[0].split('/').reverse().join('-') + ' ' + a.timestamp.split(' ')[1]));

    return { success: true, history };
  } catch (e) {
    logError('getBuyPlanHistory', e);
    return { success: false, error: e.message, history: [] };
  }
}

function _readBuyPlanConfig(ticker, market) {
  const sheet = getSheet(BUY_PLAN_SHEET.NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < BUY_PLAN_SHEET.START_ROW) return null;

  const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
  const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 16).getValues(); // ← 15→16

  for (let i = 0; i < rows.length; i++) {
    const rTicker = String(rows[i][0] || '').trim().toUpperCase();
    const rMarket = String(rows[i][1] || '').trim().toUpperCase();
    if (rTicker === ticker && (!market || rMarket === market)) {
      return {
        ticker: rTicker, market: rMarket,
        planType: String(rows[i][2] || 'price').trim().toLowerCase(),
        budget: parseFloat(rows[i][3]) || 0,
        legs: [
          { pct: parseFloat(rows[i][4]) || 0, triggerPct: parseFloat(rows[i][5]) },
          { pct: parseFloat(rows[i][6]) || 0, triggerPct: parseFloat(rows[i][7]) },
          { pct: parseFloat(rows[i][8]) || 0, triggerPct: parseFloat(rows[i][9]) }
        ].filter(l => l.pct > 0),
        dcaFreqDays: parseFloat(rows[i][10]) || null,
        dcaAmount: parseFloat(rows[i][11]) || null,
        startDate: rows[i][12] instanceof Date ? rows[i][12] : null,
        note: rows[i][13] || '',
        referencePrice: parseFloat(rows[i][14]) || null,
        direction: String(rows[i][15] || 'down').trim().toLowerCase() // ← ใหม่ default 'down' กันแผนเก่าพัง
      };
    }
  }
  return null;
}

function buildPriceBasedPlanStatus(cfg, buys, currentPrice, holdingInfo, fxRateForDisplay, market) {
  const isWholeShareOnly = (market === 'TH');
  const firstEntryPrice = buys.length ? buys[0].price : cfg.referencePrice;
  const usingReferencePrice = !buys.length && !!cfg.referencePrice;
  const direction = cfg.direction || 'down';
  const dirSign = direction === 'up' ? 1 : -1; // up: บวก (แพงขึ้น), down: ลบ (ถูกลง)

  let searchFromIdx = 0;
  const steps = cfg.legs.map((leg, i) => {
    if (i === 0) {
      const executed = buys.length > 0;
      const result = {
        legNumber: 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice: firstEntryPrice,
        status: executed ? 'done' : 'not_yet',
        executedPrice: executed ? buys[0].price : null,
        executedShares: executed && !isNaN(buys[0].shares) ? buys[0].shares : null,
        executedDate: executed ? Utilities.formatDate(new Date(buys[0].date), 'Asia/Bangkok', 'dd/MM/yyyy') : null,
        matchNote: usingReferencePrice ? 'อ้างอิงราคา ณ ตอนสร้างแผน (ยังไม่เคยถือ)' : null
      };
      if (executed) searchFromIdx = 1;
      return result;
    }

    if (firstEntryPrice === null) {
      return { legNumber: i + 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice: null,
        status: 'pending', executedPrice: null, executedShares: null, executedDate: null, matchNote: 'ไม่มีราคาอ้างอิง' };
    }

    // ── ทิศทาง down: targetPrice = ref × (1 - trigger%) → ต่ำกว่า ref
    //    ทิศทาง up:   targetPrice = ref × (1 + trigger%) → สูงกว่า ref ──
    const targetPrice = firstEntryPrice * (1 + dirSign * (Math.abs(leg.triggerPct) / 100));

    let matchedIdx = -1;
    for (let b = searchFromIdx; b < buys.length; b++) {
      // ── down: จับคู่ไม้ที่ซื้อราคา ≤ target (ย่อถึงแล้ว) / up: ซื้อราคา ≥ target (ทะลุขึ้นแล้ว) ──
      const matched = direction === 'down' ? (buys[b].price <= targetPrice) : (buys[b].price >= targetPrice);
      if (matched) { matchedIdx = b; break; }
    }

    if (matchedIdx !== -1) {
      searchFromIdx = matchedIdx + 1;
      return {
        legNumber: i + 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice,
        status: 'done', executedPrice: buys[matchedIdx].price,
        executedShares: !isNaN(buys[matchedIdx].shares) ? buys[matchedIdx].shares : null,
        executedDate: Utilities.formatDate(new Date(buys[matchedIdx].date), 'Asia/Bangkok', 'dd/MM/yyyy'),
        matchNote: null
      };
    }
    return { legNumber: i + 1, pct: leg.pct, triggerPct: leg.triggerPct, targetPrice,
      status: 'not_yet', executedPrice: null, executedShares: null, executedDate: null, matchNote: null };
  });

  let nextAssigned = false;
  steps.forEach(s => {
    if (s.status === 'not_yet') { s.status = nextAssigned ? 'pending' : 'next'; nextAssigned = true; }
  });

  if (currentPrice) {
    steps.forEach(s => {
      if (s.status === 'done' || s.targetPrice === null) return;
      s.distanceAmount = currentPrice - s.targetPrice;
      s.distancePct = (s.distanceAmount / currentPrice) * 100;
      // ── canBuyNow: down → ราคาปัจจุบันย่อถึงเป้าแล้ว / up → ราคาปัจจุบันทะลุเป้าขึ้นไปแล้ว ──
      s.canBuyNow = direction === 'down' ? (currentPrice <= s.targetPrice) : (currentPrice >= s.targetPrice);
    });
  }

  steps.forEach(s => {
    if (s.status === 'done') {
      s.actualBudgetSpent = s.executedShares !== null ? s.executedPrice * s.executedShares : cfg.budget * (s.pct / 100);
      s.budgetTHBEquivalent = (market === 'US' && fxRateForDisplay) ? s.actualBudgetSpent * fxRateForDisplay : null;
    } else if (s.targetPrice !== null) {
      const budgetForLeg = cfg.budget * (s.pct / 100);
      let estimatedShares = budgetForLeg / s.targetPrice;
      estimatedShares = isWholeShareOnly ? Math.floor(estimatedShares) : Math.round(estimatedShares * 10000) / 10000;
      s.budgetForLeg = budgetForLeg;
      s.estimatedShares = estimatedShares;
      s.budgetTHBEquivalent = (market === 'US' && fxRateForDisplay) ? budgetForLeg * fxRateForDisplay : null;

      const curShares = holdingInfo ? (parseFloat(holdingInfo.sharesRemain) || 0) : 0;
      const curTotalCost = holdingInfo ? (parseFloat(holdingInfo.totalCost) || 0) : 0;
      const curAvgCost = holdingInfo ? (parseFloat(holdingInfo.avgCost) || null) : null;
      const newTotalCost = curTotalCost + (s.targetPrice * estimatedShares);
      const newTotalShares = curShares + estimatedShares;
      const newAvgCost = newTotalShares > 0 ? newTotalCost / newTotalShares : s.targetPrice;

      s.curAvgCost = curAvgCost;
      s.newAvgCost = newAvgCost;
      s.avgCostChangeAmount = (curAvgCost !== null) ? (newAvgCost - curAvgCost) : null;
      s.avgCostChangePct = curAvgCost ? ((newAvgCost - curAvgCost) / curAvgCost) * 100 : null;
    }
  });

  return { budget: cfg.budget, firstEntryPrice, usingReferencePrice, direction, steps, note: cfg.note };
}


// ── แก้ไขราคาเป้าหมาย/% ของไม้ที่ยังไม่ซื้อ — สำหรับเทรดสั้นที่ต้องขยับตามความผันผวน
//    ไม่กระทบไม้ที่ done แล้ว (แก้ไม่ได้) และไม้ 1 (เป็นราคาอ้างอิงตายตัว แก้ไม่ได้เหมือนเดิม) ──
function editBuyPlanLeg(ticker, market, legNumber, newPrice, newPct) {
  try {
    ticker = String(ticker || '').trim().toUpperCase();
    market = String(market || '').trim().toUpperCase();
    legNumber = parseInt(legNumber, 10);
    if (legNumber === 1) return { success: false, error: 'แก้ไขไม้ 1 ไม่ได้ เพราะเป็นราคาอ้างอิงของทั้งแผน' };

    const price = parseFloat(newPrice);
    const pct = parseFloat(newPct);
    if (!price || price <= 0) return { success: false, error: 'กรุณาระบุราคาเป้าหมายที่ถูกต้อง' };
    if (!pct || pct <= 0) return { success: false, error: 'กรุณาระบุ % ที่ถูกต้อง' };

    const cfg = _readBuyPlanConfig(ticker, market);
    if (!cfg) return { success: false, error: 'ไม่พบแผนของ ' + ticker };
    if (legNumber < 1 || legNumber > cfg.legs.length) return { success: false, error: 'ไม่พบไม้ที่ ' + legNumber + ' ในแผนนี้' };
    if (!cfg.referencePrice) return { success: false, error: 'แผนนี้ไม่มีราคาอ้างอิง แก้ไขไม่ได้' };

    // ── ห้ามแก้ไม้ที่ done แล้ว ──
    const statusCheck = getBuyPlanForTicker(ticker, market);
    if (statusCheck.success) {
      const targetStep = statusCheck.steps.find(s => s.legNumber === legNumber);
      if (targetStep && targetStep.status === 'done') {
        return { success: false, error: 'ไม้ ' + legNumber + ' ซื้อไปแล้ว แก้ไขไม่ได้' };
      }
    }

    // ── VALIDATION ทิศทางเดิม: เทียบกับไม้ก่อนหน้าและไม้ถัดไป (ถ้ามี) ให้เรียงทิศทางเดียวกันตลอด ──
    const direction = cfg.direction || 'down';
    const invalid = (a, b) => direction === 'down' ? (a >= b) : (a <= b);
    const requirement = direction === 'down' ? 'ต่ำกว่า' : 'สูงกว่า';

    // เทียบกับไม้ก่อนหน้า (ไม้ 1 = referencePrice, ไม้ 2 = คำนวณจาก legs[0])
    let prevPrice;
    if (legNumber === 2) {
      prevPrice = cfg.referencePrice;
    } else {
      const dirSign = direction === 'up' ? 1 : -1;
      prevPrice = cfg.referencePrice * (1 + dirSign * (Math.abs(cfg.legs[legNumber - 2].triggerPct) / 100));
    }
    if (invalid(price, prevPrice)) {
      return { success: false, error: `ราคาใหม่ (${fmtNumServer(price)}) ต้อง${requirement}ไม้ก่อนหน้า (${fmtNumServer(prevPrice)})` };
    }

    // เทียบกับไม้ถัดไป (ถ้ามีและยังไม่ done)
    if (legNumber < cfg.legs.length) {
      const dirSign = direction === 'up' ? 1 : -1;
      const nextPrice = cfg.referencePrice * (1 + dirSign * (Math.abs(cfg.legs[legNumber].triggerPct) / 100));
      const nextInvalid = direction === 'down' ? (nextPrice >= price) : (nextPrice <= price);
      if (nextInvalid) {
        return { success: false, error: `ราคาใหม่ (${fmtNumServer(price)}) ทำให้ไม้ถัดไปผิดลำดับ — ต้อง${requirement === 'ต่ำกว่า' ? 'สูงกว่า' : 'ต่ำกว่า'}ไม้ถัดไป (${fmtNumServer(nextPrice)}) ด้วย` };
      }
    }

    // ── เช็ค % รวมไม่เกิน 100 (ไม้ 1 คำนวณจากส่วนที่เหลือเสมอ) ──
    const otherLegsExtraPct = cfg.legs.slice(1).reduce((sum, l, idx) => {
      const realIdx = idx + 2; // legs[1]=ไม้2(index1), legs[2]=ไม้3(index2) → legNumber จริง
      return realIdx === legNumber ? sum : sum + l.pct;
    }, 0);
    if (otherLegsExtraPct + pct >= 100) {
      return { success: false, error: '% รวมของไม้ 2-3 ต้องน้อยกว่า 100% (ตอนนี้ไม้อื่นรวมกัน ' + otherLegsExtraPct + '%)' };
    }

    // ── เขียนค่าใหม่กลับชีต ──
    const triggerPct = Math.abs((cfg.referencePrice - price) / cfg.referencePrice) * 100;
    const sheet = getSheet(BUY_PLAN_SHEET.NAME);
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    const numRows = lastRow - BUY_PLAN_SHEET.START_ROW + 1;
    const rows = sheet.getRange(BUY_PLAN_SHEET.START_ROW, 1, numRows, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() === ticker && String(rows[i][1]).trim().toUpperCase() === market) {
        targetRow = BUY_PLAN_SHEET.START_ROW + i; break;
      }
    }
    if (targetRow === -1) return { success: false, error: 'ไม่พบแถวแผนในชีต' };

    const pctCol = 5 + (legNumber - 1) * 2;      // ไม้2→col7(G), ไม้3→col9(I)
    const triggerCol = 6 + (legNumber - 1) * 2;   // ไม้2→col8(H), ไม้3→col10(J)
    sheet.getRange(targetRow, pctCol).setValue(pct);
    sheet.getRange(targetRow, triggerCol).setValue(_taxRound(triggerPct, 2));

    _logBuyPlanHistory(ticker, market, 'leg_edited', { legNumber, newPrice: price, newPct: pct });

    return { success: true, legNumber, newPrice: price, newPct: pct };
  } catch (e) {
    logError('editBuyPlanLeg', e);
    return { success: false, error: e.message };
  }
}

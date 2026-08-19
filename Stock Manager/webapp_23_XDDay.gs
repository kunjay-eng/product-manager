/** webapp_23_XDDay.gs
 * ดึงวันประกาศงบและวันขึ้นเครื่องหมาย XD ถัดไปจาก Yahoo Finance
 * ใช้ auth (crumb/cookie) ตัวเดียวกับที่ใช้ดึง P/E — reuse _getYahooCrumbAndCookie() เดิม
 */
function getEarningsCalendar(ticker, market) {
  try {
    const symbol = market === 'TH' ? (ticker + '.BK') : ticker;

    let auth;
    try {
      auth = _getYahooCrumbAndCookie();
    } catch (authErr) {
      return { success: false, error: 'ขอสิทธิ์เข้าถึง Yahoo Finance ไม่สำเร็จ: ' + authErr.message };
    }

    const buildUrl = crumb => 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/' + encodeURIComponent(symbol)
      + '?modules=calendarEvents&crumb=' + encodeURIComponent(crumb);

    let resp = UrlFetchApp.fetch(buildUrl(auth.crumb), { muteHttpExceptions: true, headers: { 'Cookie': auth.cookie, 'User-Agent': 'Mozilla/5.0' } });
    let code = resp.getResponseCode();
    let text = resp.getContentText();

    if (code === 401) {
      CacheService.getScriptCache().remove('yahoo_crumb_cookie_v1');
      auth = _getYahooCrumbAndCookie();
      resp = UrlFetchApp.fetch(buildUrl(auth.crumb), { muteHttpExceptions: true, headers: { 'Cookie': auth.cookie, 'User-Agent': 'Mozilla/5.0' } });
      code = resp.getResponseCode();
      text = resp.getContentText();
    }

    if (code !== 200) return { success: false, error: 'ดึงข้อมูลไม่สำเร็จ (HTTP ' + code + ')' };

    let json;
    try { json = JSON.parse(text); } catch (e) { return { success: false, error: 'ตอบกลับไม่ใช่ JSON ที่ถูกต้อง' }; }

    const result = json.quoteSummary && json.quoteSummary.result && json.quoteSummary.result[0];
    if (!result) return { success: false, error: 'ไม่พบข้อมูลของ ' + symbol + ' (หุ้นบางตัวไม่มีข้อมูลปฏิทินจาก Yahoo)' };

    const cal = result.calendarEvents || {};
    const earningsObj = cal.earnings || {};
    const earningsDates = (earningsObj.earningsDate || []).map(d => d.fmt).filter(Boolean);
    const exDivDate = cal.exDividendDate && cal.exDividendDate.fmt;
    const divDate = cal.dividendDate && cal.dividendDate.fmt;

    const today = new Date();
    const daysUntil = dateStr => dateStr ? Math.ceil((new Date(dateStr) - today) / 86400000) : null;

    const earningsDaysUntil = earningsDates.length > 0 ? daysUntil(earningsDates[0]) : null;
    const exDivDaysUntil = daysUntil(exDivDate);

    return {
      success: true, ticker,
      earningsDates, earningsDaysUntil,
      exDividendDate: exDivDate, exDividendDaysUntil: exDivDaysUntil,
      dividendDate: divDate,
      warning: (earningsDaysUntil !== null && earningsDaysUntil >= 0 && earningsDaysUntil <= 7)
        ? '⚠️ ใกล้วันประกาศงบ (อีก ' + earningsDaysUntil + ' วัน) ระวังความผันผวนสูง'
        : null
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


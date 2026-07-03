// ═══════════════════════════════════════════════════════════════
// IMARAT Finance Dashboard — Google Apps Script Backend v3
// Deploy: Extensions → Apps Script → Deploy → Manage deployments
//         → Edit (pencil icon) → Version: New version → Deploy
//         (This keeps the same URL — no need to update finance.html)
// v3.2 (FINAL): STRICT token auth ON. Firebase ID-token verification is
//     required on every data request (real auth on the data endpoint);
//     response caching (CacheService, 5 min, fresh=1 bypass); key rotated.
// NOTE: UrlFetchApp external-request scope already granted. If you ever paste
//       this into a fresh project, run any function once and approve the
//       "Connect to an external service" authorization prompt first.
// ═══════════════════════════════════════════════════════════════

var SHEET_ID         = '1457VGuwQWURarO-38eu5nFUQpKj3kYe1pVRNZ7nFS90';
var CLIENT_KEY       = 'IMARAT_FCC_2026_Xk7Qp9ZmRv4T';
var FIREBASE_API_KEY = 'AIzaSyBei-lo6GIehz3UT93xg2KTs5bU_1-4osQ'; // project imarat-sap-aff7f (public identifier)
var DATA_CACHE_KEY   = 'IMARAT_DATA_V1';
var DATA_CACHE_SEC   = 300; // 5 min — matches front-end auto-refresh
var TOKEN_CACHE_SEC  = 300; // re-verify each token at most every 5 min

// STRICT_AUTH=false: if the Firebase verification call itself is unreachable
// (API key referrer-restricted → HTTP 403, network error, quota), fall back to
// clientKey-only — your ORIGINAL security level — and return a loud warning,
// instead of locking out a legitimate signed-in user. Forged/expired tokens
// (HTTP 400) are still denied. Flip to true once authMode reads "token" to
// enforce sign-in strictly.
var STRICT_AUTH = true;

// ── Entry point ─────────────────────────────────────────────────
function doGet(e) {
  var output;
  try {
    var params = e && e.parameter ? e.parameter : {};

    if (params.clientKey !== CLIENT_KEY) {
      output = jsonOut({ status: 'error', message: 'Unauthorized' });
    } else if (params.action === 'malls') {
      // OPEN, SCOPED endpoint: returns ONLY Mall P&L data, no Firebase login.
      // Deliberately bypasses token auth so the standalone management dashboard
      // can read without sign-in. Exposes mall occupancy/rental figures ONLY —
      // never the full financial payload (balances, receivables, advances).
      var mallData = readMallsCached(params.fresh === '1');
      output = jsonOut({ status: 'ok', at: new Date().toISOString(), data: mallData });
    } else {
      var chk = authCheck(params.idToken);
      if (!chk.allow) {
        output = jsonOut({ status: 'error', message: 'Sign-in required', detail: chk.warn });
      } else {
        var data = readAllCached(params.fresh === '1');
        output = jsonOut({
          status: 'ok', at: new Date().toISOString(),
          authMode: chk.mode, authWarn: chk.warn, data: data
        });
      }
    }
  } catch (err) {
    output = jsonOut({ status: 'error', message: err.message });
  }
  return output;
}

// ── Auth check (graceful) ───────────────────────────────────────
// Returns {allow, mode:'token'|'degraded'|'deny', warn}. Prefers real token
// verification; degrades to clientKey-only only when the verification service
// is unreachable (never on a bad token).
function authCheck(idToken) {
  if (!idToken || typeof idToken !== 'string' || idToken.length < 100) {
    return { allow: false, mode: 'deny', warn: 'No/short ID token received.' };
  }
  var cache = CacheService.getScriptCache();
  var hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)
  );
  var ck = 'tok_' + hash;
  if (cache.get(ck) === 'ok') return { allow: true, mode: 'token', warn: '' };

  var resp;
  try {
    resp = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY,
      { method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true }
    );
  } catch (err) {
    return STRICT_AUTH
      ? { allow: false, mode: 'deny', warn: 'Verify fetch error: ' + err.message }
      : { allow: true, mode: 'degraded', warn: 'Token verify unreachable (' + err.message + ') — serving with clientKey only. Approve the UrlFetchApp authorization.' };
  }

  var code = resp.getResponseCode();
  if (code === 200) {
    var body = JSON.parse(resp.getContentText());
    if (body.users && body.users.length && body.users[0].disabled !== true) {
      try { cache.put(ck, 'ok', TOKEN_CACHE_SEC); } catch (x) {}
      return { allow: true, mode: 'token', warn: '' };
    }
    return { allow: false, mode: 'deny', warn: 'Token format ok but no active user.' };
  }
  if (code === 400) {
    return { allow: false, mode: 'deny', warn: 'Token rejected (400) — expired or invalid; sign in again.' };
  }
  // 403 / 429 / 5xx = key restriction, quota, or Google-side error = infra, not a bad user.
  return STRICT_AUTH
    ? { allow: false, mode: 'deny', warn: 'Token verify HTTP ' + code }
    : { allow: true, mode: 'degraded', warn: 'Token verify HTTP ' + code + ' (likely Firebase API-key restriction) — serving with clientKey only. Fix: Cloud Console - Credentials - browser key - Application restrictions - None.' };
}


// ── Mall-only reader (for the open standalone dashboard) ────────
// Parses ONLY the MALL OPS tab. Cached 5 min; fresh=1 bypasses.
function readMallsCached(forceFresh) {
  var cache = CacheService.getScriptCache();
  if (!forceFresh) {
    try { var hit = cache.get('IMARAT_MALLS_V1'); if (hit) return JSON.parse(hit); } catch (x) {}
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = parseMallOps(findSheet(ss, ['MALL OPS','MALL-OPS','Mall Ops','Mall Operations']));
  try { cache.put('IMARAT_MALLS_V1', JSON.stringify(data), 300); } catch (x) {}
  return data;
}

// ── Cached reader ────────────────────────────────────────────────
// Serves the parsed payload from CacheService when fresh (≤5 min old);
// manual refresh sends fresh=1 to force a re-read of the sheet.
// Payloads >100KB exceed the cache limit — put() failure is swallowed
// and every request simply reads the sheet directly (previous behavior).
function readAllCached(forceFresh) {
  var cache = CacheService.getScriptCache();
  if (!forceFresh) {
    try {
      var hit = cache.get(DATA_CACHE_KEY);
      if (hit) return JSON.parse(hit);
    } catch (x) {}
  }
  var data = readAll();
  try { cache.put(DATA_CACHE_KEY, JSON.stringify(data), DATA_CACHE_SEC); } catch (x) {}
  return data;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Master reader ────────────────────────────────────────────────
function readAll() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var hist = parsePeriodGrid(findSheet(ss, ['MTD HISTORICAL','MTD Historic','MTD Historical']), 'monthly');
  var curr = parsePeriodGrid(findSheet(ss, ['Current month MTD','Current Month MTD','CURRENT MONTH MTD']), 'daily');

  // Aggregate current-month daily columns into monthly totals and inject each
  // distinct month into historical (so Trends shows every month), while the
  // Daily tab shows only the latest month. Grouping by month prevents lumping
  // e.g. July's daily columns into June's total when both sit in one sheet.
  if (curr.periods.length > 0 && curr.rows.length > 0) {
    var monthOrder = [];
    var monthCols = {};
    for (var pp = 0; pp < curr.periods.length; pp++) {
      var ml = dailyPeriodToMonthLabel(curr.periods[pp]);
      if (!ml) continue;
      if (!monthCols[ml]) { monthCols[ml] = []; monthOrder.push(ml); }
      monthCols[ml].push(curr.periods[pp]);
    }
    for (var mo = 0; mo < monthOrder.length; mo++) {
      var monthLabel = monthOrder[mo];
      if (hist.periods.indexOf(monthLabel) >= 0) continue; // already a real historical column
      var cols = monthCols[monthLabel];
      for (var i = 0; i < curr.rows.length; i++) {
        var monthTotal = 0;
        for (var cc = 0; cc < cols.length; cc++) {
          monthTotal += (curr.rows[i].values[cols[cc]] || 0);
        }
        for (var j = 0; j < hist.rows.length; j++) {
          if (hist.rows[j].head === curr.rows[i].head) {
            hist.rows[j].values[monthLabel] = monthTotal;
            break;
          }
        }
      }
      hist.periods.push(monthLabel);
    }
  }

  return {
    inflow:           parseInflow(ss.getSheetByName('INFLOW DETAIL')),
    expenses:         parseMTD(ss.getSheetByName('MTD EXPENSE')),
    balances:         parseBalances(ss.getSheetByName('BALANCES')),
    pendingDetail:    parsePendingDetail(ss.getSheetByName('PENIDNG REV')),
    historical:       hist,
    currentMTD:       curr,
    rentalAdvances:   parseRentalAdvances(findSheet(ss, ['RENTAL-ADVANCES','RENTAL/ADVANCES','RENTAL ADVANCES'])),
    mallOps:          parseMallOps(findSheet(ss, ['MALL OPS','MALL-OPS','Mall Ops','Mall Operations'])),
    ibuPL:            parseIbuPL(findSheet(ss, ['IBU PL','IBU-PL','IBU P&L','IBU'])),
    productData:      parseProductData(findSheet(ss, ['PRODUCT DATA','Product Data','PRODUCT-DATA'])),
    productDetail:    parseProductDetail(findSheet(ss, ['PRODUCT DETAIL','Product Detail','PRODUCT-DETAIL'])),
    appreciation:     parseAppreciation(findSheet(ss, ['APPRECIATION','Appreciation']))
  };
}

// Convert a daily period label to a month label
// "1-Jun-2026" → "Jun-26"  |  "15-Mar-2026" → "Mar-26"
function dailyPeriodToMonthLabel(period) {
  if (!period) return null;
  var m = String(period).match(/^\d+-([A-Za-z]{3})-(\d{4})$/);
  if (!m) return null;
  var yr2 = m[2].slice(-2); // "2026" → "26"
  return m[1] + '-' + yr2;  // "Jun-26"
}

// Robust sheet finder — tries multiple name variations
function findSheet(ss, names) {
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  // Fallback: case-insensitive partial match
  var all = ss.getSheets();
  for (var j = 0; j < all.length; j++) {
    var nm = all[j].getName().toUpperCase().replace(/\s+/g, ' ').trim();
    for (var k = 0; k < names.length; k++) {
      if (nm === names[k].toUpperCase().replace(/\s+/g, ' ').trim()) return all[j];
    }
  }
  return null;
}

// ── PERIOD GRID parser (MTD HISTORICAL + Current month MTD) ───────
// Structure: Row 1 = headers [Category Code | Head | Period1 | Period2 | ...]
//            Rows 2-N = [code | head | val1 | val2 | ...]
// Periods are months (Jul-25) or dates (1-Jun-2026)
function parsePeriodGrid(sheet, type) {
  var result = { periods: [], rows: [] };
  if (!sheet) return result;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  var header = data[0];
  // Columns 2 onwards (index 2+) are periods
  var periodCols = [];
  for (var c = 2; c < header.length; c++) {
    var label = formatPeriodLabel(header[c], type);
    if (label) {
      result.periods.push(label);
      periodCols.push(c);
    }
  }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var code = String(row[0] || '').trim();
    var head = String(row[1] || '').trim();
    if (!head) continue;
    var values = {};
    for (var p = 0; p < periodCols.length; p++) {
      values[result.periods[p]] = toNum(row[periodCols[p]]);
    }
    result.rows.push({ code: code, head: head, values: values });
  }
  return result;
}

// Normalise period header — handles Date objects and strings
// type = 'monthly' → MTD HISTORICAL headers (Jul-25, Feb-26 stored as dates)
// type = 'daily'   → Current month MTD headers (1-Jun-2026 actual dates)
function formatPeriodLabel(v, type) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    if (type === 'monthly') {
      // Google Sheets misread "Jul-25" as July 25 (current year).
      // The day value (25, 26...) IS the original year suffix.
      // Reconstruct: day=25 + month=Jul → "Jul-25"
      var daySuffix = v.getDate(); // 25 → year 2025, 26 → year 2026 etc.
      var mon = Utilities.formatDate(v, Session.getScriptTimeZone(), 'MMM');
      return mon + '-' + daySuffix; // "Jul-25", "Feb-26", "Mar-26"
    } else {
      // Daily: actual date → format as "d-MMM-yyyy" (e.g. 1-Jun-2026)
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'd-MMM-yyyy');
    }
  }
  return String(v).trim();
}

// ── INFLOW DETAIL parser ─────────────────────────────────────────
// Structure: label in col A, value in col B.
// Exception: TOTAL CLEARED in col D / col E.
// Sections: "TOTAL PENDING TILL DATE" starts entity block,
//           "EXPECTED CLEARANCE FROM PENDING" starts exp block.
function parseInflow(sheet) {
  var d = {
    new_sales: 0, cleared_from_new: 0, pending_from_new: 0, total_cleared: 0,
    overall_pending: 0, cleared_from_pending: 0, pending_till_date: 0,
    entity_graana: 0, entity_a21: 0, entity_imarat: 0,
    on_hold: 0, returned: 0, chq_not_received: 0, postdated: 0,
    presented: 0, to_be_deposited: 0, cleared: 0,
    exp_a21: 0, exp_graana: 0, exp_imarat: 0, exp_total: 0,
    as_of_label: ''
  };
  if (!sheet) return d;

  var rows = sheet.getDataRange().getValues();
  var sec = '';

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var lbl = String(row[0] || '').trim().toUpperCase();
    var val = toNum(row[1]);

    if (lbl === 'NEW SALES')                                          d.new_sales = val;
    else if (lbl.indexOf('CLEARED REVENUE FROM NEW') >= 0)           d.cleared_from_new = val;
    else if (lbl.indexOf('PENDING FROM NEW') >= 0)                    d.pending_from_new = val;
    else if (lbl === 'OVERALL PENDING')                               d.overall_pending = val;
    else if (lbl.indexOf('CLEARED PAYMENT FROM PENDING') >= 0)       d.cleared_from_pending = val;
    else if (lbl.indexOf('PENDING TILL') >= 0 && lbl.indexOf('TOTAL') < 0) {
      d.pending_till_date = val;
      d.as_of_label = String(row[0]).trim();
    }
    else if (lbl === 'ON HOLD')                                       d.on_hold = val;
    else if (lbl === 'RETURNED')                                      d.returned = val;
    else if (lbl.indexOf('CHQ NOT') >= 0)                            d.chq_not_received = val;
    else if (lbl === 'POSTDATED')                                     d.postdated = val;
    else if (lbl === 'PRESENTED')                                     d.presented = val;
    else if (lbl.indexOf('TO BE DEPOSITED') >= 0)                    d.to_be_deposited = val;
    else if (lbl === 'CLEARED')                                       d.cleared = val;

    // TOTAL CLEARED sits in col D/E on the same row as empty col A
    if (row[3] && String(row[3]).toUpperCase().indexOf('TOTAL CLEARED') >= 0) {
      d.total_cleared = toNum(row[4]);
    }

    // Section transitions
    if (lbl === 'TOTAL PENDING TILL DATE')              { sec = 'ent'; continue; }
    if (lbl.indexOf('EXPECTED CLEARANCE') >= 0)         { sec = 'exp'; continue; }
    if (lbl.indexOf('TOTAL PENDING BIFURCATION') >= 0)  { sec = '';    continue; }

    if (sec === 'ent') {
      if (lbl === 'GRAANA')         d.entity_graana  = val;
      else if (lbl === 'AGENCY 21') d.entity_a21     = val;
      else if (lbl === 'IMARAT')    d.entity_imarat  = val;
    }
    if (sec === 'exp') {
      if (lbl === 'AGENCY 21')      d.exp_a21    = val;
      else if (lbl === 'GRAANA')    d.exp_graana = val;
      else if (lbl === 'IMARAT')    d.exp_imarat = val;
      else if (lbl === 'TOTAL')     { d.exp_total = val; sec = ''; }
    }
  }

  return d;
}

// ── MTD EXPENSE parser ───────────────────────────────────────────
// Row 1 = header. Cols: Date | Cat | Head | Today | MTD | EnteredBy | Month
function parseMTD(sheet) {
  var result = { expenses: [], totalToday: 0, totalMTD: 0 };
  if (!sheet) return result;

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[2]) continue;
    var head = String(r[2]).trim();
    if (!head) continue;
    var today = toNum(r[3]);
    var mtd   = toNum(r[4]);
    result.totalToday += today;
    result.totalMTD   += mtd;
    result.expenses.push({
      head:  head,
      cat:   String(r[1] || '').trim(),
      today: today,
      mtd:   mtd
    });
  }
  return result;
}

// ── BALANCES parser ──────────────────────────────────────────────
// Row 1 = header. Account rows: Date|Bank|Code|Working|Float|Ledger
// Summary rows (col A null): label in col C, value in col D
function parseBalances(sheet) {
  var result = {
    accounts: [], currentBalance: 0, cashInHand: 0, totalWorking: 0, float: 0
  };
  if (!sheet) return result;

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];

    // Summary rows — col A is empty, label in col C
    if (!r[0] && r[2]) {
      var lbl = String(r[2]).trim().toUpperCase();
      if (lbl === 'CURRENT BALANCE') {
        result.currentBalance = toNum(r[3]);
        result.float          = toNum(r[4]);
      } else if (lbl === 'CASH IN HAND') {
        result.cashInHand = toNum(r[3]);
      } else if (lbl.indexOf('TOTAL WORKING') >= 0) {
        result.totalWorking = toNum(r[3]);
      }
      continue;
    }

    // Account rows
    if (!r[0] || !r[2]) continue;
    var code = String(r[2] || '').trim();
    if (!code) continue;

    result.accounts.push({
      bank:    String(r[1] || '').trim(),
      code:    code,
      working: toNum(r[3]),
      float:   toNum(r[4]),
      ledger:  toNum(r[5])
    });
  }

  // Fallback if TOTAL WORKING BALANCE row is absent
  if (!result.totalWorking) {
    result.totalWorking = result.currentBalance + result.cashInHand;
  }

  return result;
}

// ── PENIDNG REV parser ────────────────────────────────────────────
// Row 1 = header (17 cols). Data rows 2–N.
// Cols: 0=EntryDate 1=TxnDate 2=Company 3=Office 4=Advisor 5=PortfolioLead
//       6=ZonalPM 7=RegMgr 8=Team 9=Customer 10=Status 11=Amount
//       12=ExpectedClearance 13=Remarks 14=DaysAging 15=RiskFlag 16=Month
function parsePendingDetail(sheet) {
  var result = [];
  if (!sheet) return result;

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var company = String(r[2] || '').trim();
    if (!company) continue;
    var amount = toNum(r[11]);
    if (!amount) continue;

    result.push({
      entryDate:         fmtDate(r[0]),
      company:           company,
      office:            String(r[3]  || '').trim(),
      advisor:           String(r[4]  || '').trim(),
      regMgr:            String(r[7]  || '').trim(),
      team:              String(r[8]  || '').trim(),
      customer:          String(r[9]  || '').trim(),
      status:            String(r[10] || '').trim(),
      amount:            amount,
      expectedClearance: fmtDate(r[12]),
      daysAging:         r[14] ? toNum(r[14]) : null,
      riskFlag:          String(r[15] || '').trim(),
      month:             String(r[16] || '').trim()
    });
  }
  return result;
}

// ── Utilities ────────────────────────────────────────────────────
function toNum(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  var n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function fmtDate(v) {
  if (!v) return null;
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}


// ── RENTAL-ADVANCES parser ────────────────────────────────────────
// Section 1: rows per project per month (Code|Name|Month|CC cols|CA cols|RT cols|NetRent)
// Section 2: cumulative advances (ProjectName|Amount|AsOf)
function parseRentalAdvances(sheet) {
  var result = {
    monthly:        [],   // [{month, projects:[{code,name,count,gross,wht,deduction,netRent}]}]
    latestMonth:    null,
    cumulative:     [],   // [{name,amount,asOf}]
    cumulativeAsOf: null,
    grandTotal:     0,
    cumApproved:    0,    // advances = approved + cleared
    cumCleared:     0
  };
  if (!sheet) return result;

  var data = sheet.getDataRange().getValues();
  var monthMap   = {};
  var inSection2 = false;
  var defaultMonth = 'May-26';

  for (var r = 0; r < data.length; r++) {
    var row  = data[r];
    var colA = String(row[0] || '').trim();

    // Detect Section 2 boundary
    if (colA.toUpperCase().indexOf('SECTION 2') >= 0 || colA.toUpperCase().indexOf('CUMULATIVE') >= 0) {
      inSection2 = true; continue;
    }

    // ── Section 1 — Monthly Rent (Project, Count, Gross, WHT, Deduction, Net) ──
    if (!inSection2) {
      if (!colA || colA === 'Project' || colA === 'TOTAL' || colA === 'Total') continue;
      var count = toNum(row[1]);
      var gross = toNum(row[2]);
      // Skip header / non-data rows (a project row must have a count or gross)
      if (count === 0 && gross === 0) continue;
      var wht   = toNum(row[3]);
      var ded   = toNum(row[4]);
      var net   = toNum(row[5]);
      if (net === 0 && gross > 0) net = gross - wht - ded; // fallback compute
      if (!monthMap[defaultMonth]) monthMap[defaultMonth] = [];
      monthMap[defaultMonth].push({
        code:      colA,
        name:      colA,
        month:     defaultMonth,
        count:     count,
        gross:     gross,
        wht:       wht,
        deduction: ded,
        netRent:   net
      });
    }

    // ── Section 2 — cumulative advances (Project, Approved, Cleared, Total, AsOf) ──
    if (inSection2) {
      var colB = row[1]; // Approved amount
      if (!colA || colA === 'Project Name (Full)' || colA === 'Project Name') continue;
      var approved = toNum(colB);
      var cleared  = toNum(row[2]);
      var totalAmt = toNum(row[3]);  // Total Amount (PKR) = Approved + Cleared
      var asOf     = String(row[4] || '').trim();
      if (colA.toUpperCase() === 'GRAND TOTAL') {
        result.grandTotal = totalAmt || (approved + cleared);
        result.cumApproved = approved;
        result.cumCleared = cleared;
        continue;
      }
      if (!totalAmt && !approved) continue;
      if (!result.cumulativeAsOf && asOf) result.cumulativeAsOf = asOf;
      result.cumulative.push({
        name: colA,
        amount: totalAmt || (approved + cleared),
        approved: approved,
        cleared: cleared,
        asOf: asOf
      });
    }
  }

  var months = Object.keys(monthMap);
  result.monthly = months.map(function(m) { return { month: m, projects: monthMap[m] }; });
  result.latestMonth = months.length ? months[months.length - 1] : null;
  return result;
}


// ── MALL OPS parser v2 ────────────────────────────────────────────
// Sign convention: OpExp=NEGATIVE, CAM=POSITIVE, FCExp=NEGATIVE, GymExp=NEGATIVE
// Cols: A=Mall B=Month C=Period D=RentalIncome E=OpExp F=CAM G=NOI(auto)
//       H=FCRev I=FCExp J=GymRev K=GymExp L=LeasedSqFt M=TotalSqFt N=RatePSQFT(auto)
//       O=PropertyValuation P=TargetMonthlyRental(auto) Q=TargetYield% R=ActualYield%(auto)
// Coerce a Month cell to the "MMM-YY" label the dashboard expects.
// If a data-entry person types a real date (cell becomes a Date object,
// e.g. "Tue May 26 2026 …"), convert it to "May-26" instead of showing
// the raw timestamp. Text values like "May-26" pass through untouched.
function normalizeMonthLabel(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    var mons = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return mons[val.getMonth()] + '-' + String(val.getFullYear()).slice(-2);
  }
  return String(val || '').trim();
}

function parseMallOps(sheet) {
  var result = { rows: [], months: [], malls: ['AOM','IBM'] };
  if (!sheet) return result;
  var data = sheet.getDataRange().getValues();
  var monthSet = {};
  for (var r = 4; r < data.length; r++) {
    var row = data[r];
    var mall = String(row[0] || '').trim().toUpperCase();
    var month = normalizeMonthLabel(row[1]);
    if (!mall || !month || mall === 'MALL') continue;
    if (mall !== 'AOM' && mall !== 'IBM') continue;
    var ri   = toNum(row[3]);
    var oe   = toNum(row[4]);  // negative
    var cam  = toNum(row[5]);  // positive
    var noi  = toNum(row[6]);  // auto from sheet
    var fcr  = toNum(row[7]);
    var fce  = toNum(row[8]);  // negative
    var gr   = toNum(row[9]);
    var ge   = toNum(row[10]); // negative
    var lsf  = toNum(row[11]);
    var tsf  = toNum(row[12]);
    var psf  = lsf > 0 ? ri / lsf : 0;
    var occ  = tsf > 0 ? (lsf / tsf * 100) : 0;
    // Rental yield columns (new in v2)
    var propVal     = row.length > 14 ? toNum(row[14]) : 0;
    var targetRent  = row.length > 15 ? toNum(row[15]) : (propVal > 0 ? propVal * 0.05 / 12 : 0);
    var targetYield = row.length > 16 ? toNum(row[16]) : 5;
    var actualYield = row.length > 17 ? toNum(row[17]) : (propVal > 0 ? ri * 12 / propVal * 100 : 0);
    result.rows.push({
      mall: mall, month: month,
      rentalIncome: ri,
      opExp: Math.abs(oe),
      cam: cam,
      noi: noi,
      fcRevenue: fcr, fcExpense: Math.abs(fce), fcNoi: fcr + fce,
      gymRevenue: gr, gymExpense: Math.abs(ge), gymNoi: gr + ge,
      leasedSqFt: lsf, totalSqFt: tsf,
      ratePSQFT: Math.round(psf), occupancy: parseFloat(occ.toFixed(1)),
      propValuation: propVal,
      targetMonthlyRent: targetRent,
      targetYieldPct: targetYield,
      actualYieldPct: parseFloat(actualYield.toFixed(2))
    });
    if (!monthSet[month]) { monthSet[month] = true; result.months.push(month); }
  }
  return result;
}

// ── IBU PL parser v2 ──────────────────────────────────────────────
// Cols: A=IBU B=Month C=Emp D=InflowRental E=InflowInvest F=InflowInstall
//       G=TotalInflows(auto) H=Revenue I=GP J=GP%(auto)
//       K=StaffCosts L=MktgEvents M=Occupancy N=Operations O=TotalExp(auto)
//       P=EBITDA(auto) Q=EBITDA%(auto) R=DA S=Interest T=EBT(auto)
//       U=BudgetInflows V=BudgetGP W=BudgetEBITDA(auto)
function parseIbuPL(sheet) {
  var result = { rows: [], months: [], ibus: [] };
  if (!sheet) return result;
  var data = sheet.getDataRange().getValues();
  var monthSet = {}, ibuSet = {};
  for (var r = 4; r < data.length; r++) {
    var row = data[r];
    var ibu   = String(row[0] || '').trim();
    var month = String(row[1] || '').trim();
    if (!ibu || !month || ibu === 'IBU Name' || ibu === 'GROUP TOTAL') continue;
    if (month === 'Month') continue;
    var emp  = toNum(row[2]);
    var ir   = toNum(row[3]), ii = toNum(row[4]), iin = toNum(row[5]);
    var ti   = ir + ii + iin;
    var rev  = toNum(row[7]);
    var gp   = toNum(row[8]);
    var gpp  = rev > 0 ? gp / rev : 0;
    var staff= toNum(row[10]), mktg = toNum(row[11]);
    var occ  = toNum(row[12]), ops  = toNum(row[13]);
    var tExp = staff + mktg + occ + ops;
    var ebitda = gp - tExp;
    var ebitdaPct = rev > 0 ? ebitda / rev : 0;
    var da   = toNum(row[17]), interest = toNum(row[18]);
    var ebt  = ebitda - da - interest;
    var bi   = toNum(row[20]), bgp = toNum(row[21]);
    result.rows.push({
      ibu: ibu, month: month, employees: emp,
      inflowRental: ir, inflowInvest: ii, inflowInstall: iin,
      totalInflows: ti, revenue: rev,
      grossProfit: gp, gpPct: gpp,
      staffCosts: staff, mktgEvents: mktg,
      occupancy: occ, operations: ops,
      totalExpenses: tExp,
      ebitda: ebitda, ebitdaPct: ebitdaPct,
      da: da, interest: interest, ebt: ebt,
      budgetInflows: bi, budgetGP: bgp
    });
    if (!monthSet[month]) { monthSet[month] = true; result.months.push(month); }
    if (!ibuSet[ibu])     { ibuSet[ibu]   = true; result.ibus.push(ibu);   }
  }
  return result;
}


// ── PRODUCT DATA parser ───────────────────────────────────────────
// Cols: A=Month B=Product C=Approved D=Cleared E=InProgress F=Pending G=Total
function parseProductData(sheet) {
  var result = { rows: [], months: [], products: [] };
  if (!sheet) return result;
  var data = sheet.getDataRange().getValues();
  var monthSet = {}, prodSet = {};
  for (var r = 3; r < data.length; r++) {
    var row = data[r];
    var month   = String(row[0] || '').trim();
    var product = String(row[1] || '').trim();
    if (!month || !product || month === 'Month' || product === 'TOTAL' || product === 'All Products') continue;
    var appr = toNum(row[2]);
    var clr  = toNum(row[3]);
    var inp  = toNum(row[4]);
    var pend = toNum(row[5]);
    var advances = appr + clr;              // realized money only
    var pipeline = appr + clr + inp + pend; // full pipeline
    result.rows.push({
      month: month, product: product,
      approved: appr, cleared: clr,
      inProgress: inp, pending: pend,
      advances: advances,
      total: advances,    // "total" now means advances (Appr+Clear)
      pipeline: pipeline
    });
    if (!monthSet[month])   { monthSet[month]   = true; result.months.push(month); }
    if (!prodSet[product])  { prodSet[product]  = true; result.products.push(product); }
  }
  return result;
}

// ── APPRECIATION parser ────────────────────────────────────────────
// Cols: A=Month B=Project C=Count D=GrossAppreciation E=Deductions F=Net(auto)
function parseAppreciation(sheet) {
  var result = { rows: [], months: [], projects: [] };
  if (!sheet) return result;
  var data = sheet.getDataRange().getValues();
  var monthSet = {}, projSet = {};
  for (var r = 3; r < data.length; r++) {
    var row = data[r];
    var month   = String(row[0] || '').trim();
    var project = String(row[1] || '').trim();
    if (!month || !project || month === 'Month' || project === 'TOTAL' || project === 'All Projects') continue;
    var count = toNum(row[2]);
    var gross = toNum(row[3]);
    var ded   = toNum(row[4]);
    var net   = gross - ded;
    result.rows.push({
      month: month, project: project,
      count: count, gross: gross,
      deductions: ded, net: net,
      netPct: gross > 0 ? parseFloat((net / gross * 100).toFixed(1)) : 0
    });
    if (!monthSet[month])   { monthSet[month]   = true; result.months.push(month); }
    if (!projSet[project])  { projSet[project]  = true; result.projects.push(project); }
  }
  return result;
}


// ── PRODUCT DETAIL parser ─────────────────────────────────────────
// Cols: A=Month B=Product C=Frequency D=PaymentType E=Amount F=Cleared
function parseProductDetail(sheet) {
  var result = { rows: [], months: [], products: [], frequencies: [], paymentTypes: [] };
  if (!sheet) return result;
  var data = sheet.getDataRange().getValues();
  var mSet = {}, pSet = {}, fSet = {}, tSet = {};
  for (var r = 3; r < data.length; r++) {
    var row = data[r];
    var month   = String(row[0] || '').trim();
    var product = String(row[1] || '').trim();
    var freq    = String(row[2] || '').trim();
    var ptype   = String(row[3] || '').trim();
    if (!month || !product) continue;
    if (month === 'Month' || product === 'GRAND TOTAL' || product.indexOf('All product') === 0) continue;
    var amount  = toNum(row[4]);
    var cleared = toNum(row[5]);
    if (amount === 0 && cleared === 0) continue;
    result.rows.push({
      month: month, product: product,
      frequency: freq || '\u2014',
      paymentType: ptype || '(blank)',
      amount: amount, cleared: cleared
    });
    if (!mSet[month])   { mSet[month] = true; result.months.push(month); }
    if (!pSet[product]) { pSet[product] = true; result.products.push(product); }
    if (freq && !fSet[freq]) { fSet[freq] = true; result.frequencies.push(freq); }
    if (ptype && !tSet[ptype]) { tSet[ptype] = true; result.paymentTypes.push(ptype); }
  }
  return result;
}

// ── Local test: run from Apps Script editor to verify ───────────
function testRead() {
  var result = readAll();
  Logger.log('=== IMARAT Finance Apps Script Test ===');
  Logger.log('Inflow pending_till_date: ' + result.inflow.pending_till_date);
  Logger.log('Inflow entity_graana: '     + result.inflow.entity_graana);
  Logger.log('Inflow entity_a21: '        + result.inflow.entity_a21);
  Logger.log('Balances totalWorking: '    + result.balances.totalWorking);
  Logger.log('Balances accounts: '        + result.balances.accounts.length);
  Logger.log('Expenses totalMTD: '        + result.expenses.totalMTD);
  Logger.log('Expenses rows: '            + result.expenses.expenses.length);
  Logger.log('Pending rows: '             + result.pendingDetail.length);
}

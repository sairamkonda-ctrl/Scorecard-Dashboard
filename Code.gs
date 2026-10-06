/*************************************************************
 * mediamint — Operations KPI Dashboard  (Google Apps Script)
 * ===========================================================
 * SETUP
 *  1. Spreadsheet tabs (exact names):
 *       Queue Tracker · Rejection Categories · Team List
 *       Utilization · External Errors
 *  2. Extensions ▸ Apps Script ▸ paste this into Code.gs.
 *  3. File ▸ + ▸ HTML ▸ name it "Index" ▸ paste Index.html.
 *  4. Deploy ▸ New deployment ▸ Web app
 *       Execute as: Me   |   Who has access: <your domain>
 *
 * ACCESS: emails in LEAD_EMAILS see everyone; everyone else
 * only ever receives their own rows from the server.
 *
 * IDENTITY STITCHING: if a person's rows use a slightly
 * different email in one tab than another, the Name column on
 * Queue Tracker / Rejection Categories is used to re-link them
 * to the correct person automatically. Run findSplits() to see
 * what got stitched, or diagnose('name') to inspect raw emails.
 *************************************************************/

const SHEET_ID = '';
const LEAD_EMAILS = ['sairam.konda@mediamint.com'];
// Only this exact person sees / can run the "Export to Google Sheet" button.
const EXPORT_LEAD_EMAIL = 'sairam.konda@mediamint.com';
const CAPACITY_HOURS = 7.5;

// Nigeria-dedicated report (see the "NIGERIA-DEDICATED REPORT" section near
// sendNigeriaDigestEmail): must match the "Country" column (Col C) on Team List.
const NIGERIA_COUNTRY_VALUE = 'Nigeria';
const NIGERIA_REPORT_TO = ['yusuf.aderinto@mediamint.com', 'samuel.samuel@mediamint.com'];
const NIGERIA_REPORT_CC = ['avinash.vellore@mediamint.com', 'pavan.davuluri@mediamint.com', 'sharath.upadhyay@mediamint.com', 'sairam.konda@mediamint.com'];

// Display name shown as the email sender (instead of the raw script-owner email).
const EMAIL_SENDER_NAME = 'EMEA Social Scorecard System';
// Public web-app URL of the dashboard. Leave '' to auto-detect the deployed URL.
const DASHBOARD_URL = '';
function dashboardUrl_(){ if (DASHBOARD_URL) return DASHBOARD_URL; try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

// AHT = production time per handled item.
//   'taskcount'   -> per task (Assigned By + Trafficking + Live QC)  (default)
//   'trafficking' -> per trafficked item only
//   'liveqc'      -> per Live QC item only
// 'taskcount' is the safe default: people who only do QC (zero trafficking)
// still get a real AHT instead of 0:00.
const AHT_BASIS = 'taskcount';

// Manual override if two spellings can't be auto-linked (no Name to go on).
// 'wrong.email@x.com': 'correct.email@x.com'
const EMAIL_ALIASES = {
};

const TABS = {
  queue: 'Queue Tracker', rej: 'Rejection Categories', team: 'Team List',
  util: 'Utilization', ext: 'External Errors', score: 'Scorecard',
  emailStatus: 'Email Status',
  checklist: 'Checklist Discrepancies', pkt: 'PKT',
  selfdev: 'Self Developement', comp: 'Compliance', pa: 'Process Adherence',
  leaves: 'Leaves' // rename to match the actual tab name if it differs
};

/* Corporate Scorecard — KPI weightages (from the Scorecard rubric, sum = 1.0) */
var KPI_WEIGHTS = {
  launch:.15, qa:.15, checklist:.10, pkt:.10,   // Core Delivery Excellence (0.5)
  pa:.15, util:.15,                              // Client & Business Impact (0.3)
  training:.10,                                  // Self Development (0.1)
  compliance:.10                                 // Compliance (0.1)
};

/* ---------- web entry ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('EMEA Social - Operations KPI Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getViewerEmail() {
  try {
    var e = Session.getActiveUser().getEmail();
    if (!e) { try { e = Session.getEffectiveUser().getEmail(); } catch (e2) {} }
    e = normEmail_(e || '');
    if (!e && LEAD_EMAILS.length) e = normEmail_(LEAD_EMAILS[0]);   // fallback for the lead
    return e;
  } catch (e) { return LEAD_EMAILS.length ? normEmail_(LEAD_EMAILS[0]) : ''; }
}

/* Minimal authorization + delivery test. Run this once from the editor:
 * it does NOTHING but send a single plain email to you. Running it forces
 * Google to show the Gmail authorization prompt. Approve it, then check inbox. */
function sendTestPing() {
  var to = getViewerEmail();
  Logger.log('Sending ping to: ' + to);
  MailApp.sendEmail(to, 'KPI Dashboard — test ping ✅',
    'If you can read this, email sending works and the Gmail scope is authorized.\n\nNext: run testDaily() to preview the real scorecard email.');
  Logger.log('Ping sent. Check inbox (and Spam / Promotions) for: ' + to);
  return 'Ping sent to ' + to;
}

/* ---------- helpers ---------- */
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function pad_(x){ x = String(x); return x.length < 2 ? '0' + x : x; }
function num_(v){ var n = Number(v); return isNaN(n) ? 0 : n; }
function r4_(v){ return Math.round(num_(v) * 1e4) / 1e4; }

// emails contain no whitespace -> strip spaces / NBSP / zero-width / BOM, lowercase
function normEmail_(v){ return String(v == null ? '' : v).toLowerCase().replace(/[\s\u00a0\u200b\u200c\u200d\ufeff]/g, ''); }
function normName_(v){ return String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim(); }
// loose key: drop spaces + collapse repeated letters ("vaishhnavi"->"vaishnavi")
function looseName_(v){ return normName_(v).replace(/\s/g, '').replace(/(.)\1+/g, '$1'); }

function toISO_(v, tz) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  var s = String(v).trim(); if (!s) return null;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (m) return m[3] + '-' + pad_(m[1]) + '-' + pad_(m[2]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  var d = new Date(s); if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return null;
}
function readSheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) return { idx: {}, rows: [] };
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { idx: {}, rows: [] };
  var idx = {}; values[0].forEach(function(h, j){ idx[String(h).trim()] = j; });
  return { idx: idx, rows: values.slice(1) };
}
function col_(sheet, row, name, alt) {
  var j = sheet.idx[name];
  if (j === undefined && alt) for (var i = 0; i < alt.length; i++) if (sheet.idx[alt[i]] !== undefined) { j = sheet.idx[alt[i]]; break; }
  return j === undefined ? '' : row[j];
}

/* ---------- identity resolution ----------
 * Returns { canon(email) } that maps any email string to the one
 * canonical email for that person, using Team List + Name columns. */
function buildIdentity_(qs, rs, ts) {
  var teamEmails = {}, nameToEmail = {}, looseMap = {};
  ts.rows.forEach(function(r){
    var e = normEmail_(col_(ts, r, 'Email')); if (!e) return;
    teamEmails[e] = true;
    var n = normName_(col_(ts, r, 'Name'));
    if (n && !(n in nameToEmail)) nameToEmail[n] = e;
    if (n) { var lk = looseName_(n); (looseMap[lk] = looseMap[lk] || {})[e] = true; }
  });
  function emailForName(nameRaw){
    var n = normName_(nameRaw); if (!n) return null;
    if (nameToEmail[n]) return nameToEmail[n];
    var s = looseMap[looseName_(n)];
    if (s) { var ks = Object.keys(s); if (ks.length === 1) return ks[0]; }   // only if unambiguous
    return null;
  }
  // learn: an email NOT in Team List whose Name points to a Team List email
  var learned = {};
  function learn(sheet, emailHdr, nameHdr){
    sheet.rows.forEach(function(r){
      var e = normEmail_(col_(sheet, r, emailHdr)); if (!e || teamEmails[e]) return;
      var cE = emailForName(col_(sheet, r, nameHdr));
      if (cE && cE !== e && !learned[e]) learned[e] = cE;
    });
  }
  learn(qs, 'Email', 'Name');
  learn(rs, 'Email', 'Name (Filtered)');
  var aliasMap = {};
  Object.keys(learned).forEach(function(k){ aliasMap[k] = learned[k]; });
  Object.keys(EMAIL_ALIASES).forEach(function(k){ aliasMap[normEmail_(k)] = normEmail_(EMAIL_ALIASES[k]); }); // manual wins
  return {
    canon: function(e){ e = normEmail_(e); return aliasMap[e] || e; },
    teamEmails: teamEmails, aliasMap: aliasMap
  };
}

/* ---------- main feed ---------- */
function getData() {
  var ss = ss_(); var tz = ss.getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var qs = readSheet_(TABS.queue), rs = readSheet_(TABS.rej), ts = readSheet_(TABS.team),
      us = readSheet_(TABS.util), es = readSheet_(TABS.ext), sc = readSheet_(TABS.score),
      cs = readSheet_(TABS.checklist), ps = readSheet_(TABS.pkt),
      pas = readSheet_(TABS.pa), sds = readSheet_(TABS.selfdev), cps = readSheet_(TABS.comp),
      lvs = readSheet_(TABS.leaves);
  var ID = buildIdentity_(qs, rs, ts);
  var canon = ID.canon;

  // names per canonical email (prefer Team List spelling, then Queue)
  var nameByEmail = {};
  ts.rows.forEach(function(r){ var e = canon(col_(ts, r, 'Email')), n = String(col_(ts, r, 'Name')||'').trim(); if (e && n && !(e in nameByEmail)) nameByEmail[e] = n; });
  qs.rows.forEach(function(r){ var e = canon(col_(qs, r, 'Email')), n = String(col_(qs, r, 'Name')||'').trim(); if (e && n && !(e in nameByEmail)) nameByEmail[e] = n; });

  var set = {};
  function add(e){ e = canon(e); if (e) set[e] = true; }
  qs.rows.forEach(function(r){ add(col_(qs, r, 'Email')); });
  rs.rows.forEach(function(r){ add(col_(rs, r, 'Email')); });
  us.rows.forEach(function(r){ add(col_(us, r, 'Email')); });
  es.rows.forEach(function(r){ add(col_(es, r, 'Email')); });
  var emails = Object.keys(set).sort();
  var idx = {}; emails.forEach(function(e, i){ idx[e] = i; });
  var names = emails.map(function(e){ return nameByEmail[e] || e.split('@')[0]; });

  // country per person (Team List, Col C = "Country") — drives the dashboard's country dropdown.
  // Aligned with emails/names, so countries[i] belongs to the person at index i.
  var countryByEmail = {};
  ts.rows.forEach(function(r){
    var e = canon(col_(ts, r, 'Email')), c = String(col_(ts, r, 'Country', ['Region']) || '').trim();
    if (e && c && !(e in countryByEmail)) countryByEmail[e] = c;
  });
  var countries = emails.map(function(e){ return countryByEmail[e] || ''; });

  var queue = [];
  qs.rows.forEach(function(r){
    var e = canon(col_(qs, r, 'Email')); if (!e) return; var d = toISO_(col_(qs, r, 'Date'), tz); if (!d) return;
    queue.push([d, idx[e], num_(col_(qs, r, 'Assigned By')), num_(col_(qs, r, 'Trafficking')),
      num_(col_(qs, r, 'Live QC')), num_(col_(qs, r, 'Rejections')), num_(col_(qs, r, 'CIL Count'))]);
  });
  var rej = [];
  rs.rows.forEach(function(r){
    var e = canon(col_(rs, r, 'Email')); if (!e) return; var d = toISO_(col_(rs, r, 'Date'), tz); if (!d) return;
    rej.push([d, idx[e], String(col_(rs, r, 'Rejection Category')||'').trim(), num_(col_(rs, r, 'Count'))]);
  });
  var util = [];
  us.rows.forEach(function(r){
    var e = canon(col_(us, r, 'Email')); if (!e) return; var d = toISO_(col_(us, r, 'Date [PT]', ['Date','Date[PT]']), tz); if (!d) return;
    util.push([d, idx[e], r4_(col_(us, r, 'Non Production')), r4_(col_(us, r, 'Production')), r4_(col_(us, r, 'Total'))]);
  });
  var ext = [];
  es.rows.forEach(function(r){
    var e = canon(col_(es, r, 'Email')); if (!e) return; var d = toISO_(col_(es, r, 'Date'), tz); if (!d) return;
    ext.push([d, idx[e], String(col_(es, r, 'Issue Type')||'').trim(), String(col_(es, r, 'Task Name')||'').trim(),
      String(col_(es, r, 'Link')||'').trim(), String(col_(es, r, 'Issue')||'').trim()]);
  });
  var score = [];
  sc.rows.forEach(function(r){
    var e = canon(col_(sc, r, 'Email')); if (!e) return; var d = toISO_(col_(sc, r, 'Date'), tz);
    if (!d) return; // Need a date
    var pr = String(col_(sc, r, 'P Rating') || 'N/A').trim();
    var pgr = String(col_(sc, r, 'PG Rating') || 'N/A').trim();
    // Defaulting Process Adherence, Compliance, and Self Development to 100% as requested.
    var pa = String(col_(sc, r, 'Process Adherence') || '100%').trim();
    var comp = String(col_(sc, r, 'Compliance') || '100%').trim();
    var sd = String(col_(sc, r, 'Self Development') || '100%').trim();
    score.push([d, idx[e], pr, pgr, pa, comp, sd]);
  });

  // name -> idx (for tabs that key on a person's name rather than email)
  var nameToIdx = {};
  emails.forEach(function(e, i){ var n = normName_(names[i]); if (n) nameToIdx[n] = i; });
  function idxForRow_(sheet, r) {
    var e = canon(col_(sheet, r, 'Email', ['Email ID']));
    if (e && idx.hasOwnProperty(e)) return idx[e];
    var nm = normName_(col_(sheet, r, 'Operator Name', ['Missed By', 'Name', 'Name (Filtered)', 'Employee Name']));
    if (nm && nameToIdx.hasOwnProperty(nm)) return nameToIdx[nm];
    if (nm) { var lk = looseName_(nm); for (var i = 0; i < names.length; i++) if (looseName_(names[i]) === lk) return i; }
    return -1;
  }

  // Checklist Discrepancies -> one row per miss: [iso, idx]
  var chk = [];
  cs.rows.forEach(function(r){
    var i = idxForRow_(cs, r); if (i < 0) return;
    var d = toISO_(col_(cs, r, 'Run Time', ['Date', 'Date [PT]']), tz) || '';
    var cat = String(col_(cs, r, 'Checklist', ['Discrepancy', 'Category', 'Type', 'Check', 'Parameter', 'Section', 'Item'])||'').trim();
    chk.push([d, i, cat]);
  });

  // PKT -> average score (%) per person, indexed by idx.
  // A person who is active (has queue/util/ext activity) in a month where the
  // team ran a PKT assessment but who has no PKT record that month counts as
  // 0 for that month, so the average is taken over every expected month, not
  // just the months they actually submitted (e.g. May+Jun submitted, Jul
  // missed -> average of 3 months, not 2).
  var activeMonths_ = {}; // idx -> { 'YYYY-MM': true }
  function markActiveMonths_(arr){
    arr.forEach(function(r){
      var mo = r[0] ? String(r[0]).slice(0, 7) : null; if (!mo) return;
      (activeMonths_[r[1]] = activeMonths_[r[1]] || {})[mo] = true;
    });
  }
  markActiveMonths_(queue); markActiveMonths_(util); markActiveMonths_(ext);

  var pktByMonth_ = {}, pktMonthsSet_ = {}, pktRows = [];
  ps.rows.forEach(function(r){
    var i = idxForRow_(ps, r); if (i < 0) return;
    var v = num_(col_(ps, r, 'PKT Score', ['Score']));
    if (!v) return;
    if (v <= 1) v *= 100;
    var d = toISO_(col_(ps, r, 'Date', ['Month', 'Assessment Date', 'Test Date', 'Date [PT]']), tz);
    var mo = d ? d.slice(0, 7) : null;
    if (mo) {
      pktMonthsSet_[mo] = true;
      var rec = (pktByMonth_[i] = pktByMonth_[i] || {});
      rec[mo] = rec[mo] || { sum: 0, n: 0 };
      rec[mo].sum += v; rec[mo].n += 1;
    }
    pktRows.push([d || '', i, r4_(v)]);
  });
  var pktMonths_ = Object.keys(pktMonthsSet_).sort();
  var pktByIdx = emails.map(function(e, i){
    if (!pktMonths_.length) return null;
    // Only count months the person was active in (or actually has a PKT score for) —
    // avoids zeroing out months before someone joined the team.
    var months = pktMonths_.filter(function(mo){
      return (activeMonths_[i] && activeMonths_[i][mo]) || (pktByMonth_[i] && pktByMonth_[i][mo]);
    });
    if (!months.length) return null;
    var total = 0;
    months.forEach(function(mo){
      var rec = pktByMonth_[i] && pktByMonth_[i][mo];
      total += rec ? (rec.sum / rec.n) : 0;
    });
    return r4_(total / months.length);
  });
  // Process Adherence (CIL details) -> [iso, idx, missedBy, columnsMissed, link]
  var paRows = [];
  pas.rows.forEach(function(r){
    var i = idxForRow_(pas, r); if (i < 0) return;
    var d = toISO_(col_(pas, r, 'Run Time', ['Date', 'Run time', 'Date [PT]']), tz) || '';
    var by = String(col_(pas, r, 'Missed By', ['Miss Type', 'Type']) || '').trim();
    var cols = String(col_(pas, r, 'Columns Missed', ['Column Missed', 'Missed Columns', 'Details']) || '').trim();
    var link = String(col_(pas, r, 'Socialite link', ['Link', 'Socialite Link', 'URL']) || '').trim();
    paRows.push([d, i, by || 'Other', cols || '—', link]);
  });

  // Self Development / Training-Certification -> [iso, idx, course, deadlineISO, delayDays|null]
  // delayDays = Completed - Deadline (negative/zero = on time). null = not completed yet.
  var selfdevRows = [];
  sds.rows.forEach(function(r){
    var i = idxForRow_(sds, r); if (i < 0) return;
    var course = String(col_(sds, r, 'Course Name', ['Course', 'Training', 'Name']) || 'Course').trim();
    var dl = toISO_(col_(sds, r, 'Deadline Date', ['Deadline', 'Target Date', 'Due Date']), tz);
    var cp = toISO_(col_(sds, r, 'Completed Date', ['Completed', 'Completion Date', 'Date']), tz);
    var delay = (dl && cp) ? Math.round((new Date(cp) - new Date(dl)) / 86400000) : null;
    selfdevRows.push([cp || dl || '', i, course, dl || '', delay]);
  });
  // training rating per person: avg of per-course bands (on-time=5, 1 day late=4 … 4+=1)
  var trnSum={}, trnN={};
  selfdevRows.forEach(function(r){ var b=trainBandDelay_(r[4]); trnSum[r[1]]=(trnSum[r[1]]||0)+b; trnN[r[1]]=(trnN[r[1]]||0)+1; });
  var trainByIdx = emails.map(function(e,i){ return trnN[i] ? Math.round(trnSum[i]/trnN[i]) : null; });

  // Compliance -> [iso, idx, type, details, isDeviation(0/1)]
  var compRows = [];
  cps.rows.forEach(function(r){
    var i = idxForRow_(cps, r); if (i < 0) return;
    var d = toISO_(col_(cps, r, 'Timestamp', ['Date', 'Run Time', 'Date [PT]']), tz) || '';
    var typ = String(col_(cps, r, 'Compliance Type', ['Type', 'Category']) || '').trim();
    var det = String(col_(cps, r, 'Details', ['Detail', 'Remarks', 'Comment']) || '').trim();
    // deviation if the note implies a delay/violation (e.g. "late by approx 8 mins")
    var low = det.toLowerCase();
    var lateM = low.match(/late by[^0-9]*([0-9]+)\s*min/);
    var dev = lateM ? (parseInt(lateM[1],10) > 0 ? 1 : 0)
                    : (/no delay|no deviation|0\s*min|on time|compliant/.test(low) ? 0 : (low ? 1 : 0));
    compRows.push([d, i, typ || 'Compliance', det, dev]);
  });
  // compliance rating per person: COUNT of logged compliance entries (fewer is better)
  //   0 -> 5 · 1 -> 4 · 2 -> 3 · 3 -> 2 · >3 -> 1   (compBand_ is the single source of truth)
  var cmpN={};
  compRows.forEach(function(r){ cmpN[r[1]]=(cmpN[r[1]]||0)+1; });
  var compByIdx = emails.map(function(e,i){ return compBand_(cmpN[i]||0); });

  // Leaves -> one row per individual date within a leave request (a single request
  // can list several dates, e.g. "May 28 (Full Day), May 25 (Full Day)" or
  // "Jun 19, 2026 (Full Day)||Jun 23, 2026 (Full Day)"):
  // [iso, idx, leaveType, dayFraction, purpose, status]
  var MONTH_IDX_ = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
  var leaveDateRe_ = /([A-Za-z]+)\s+(\d{1,2})(?:,\s*(\d{4}))?\s*\(([^)]*)\)/g;
  var nowYear_ = new Date().getFullYear();
  var leaveRows = [];
  lvs.rows.forEach(function(r){
    var i = idxForRow_(lvs, r); if (i < 0) return;
    var datesStr = String(col_(lvs, r, 'Dates Requested', ['Dates', 'Date']) || '');
    var leaveType = String(col_(lvs, r, 'Leave Type', ['Type']) || '').trim();
    var purpose = String(col_(lvs, r, 'Purpose', ['Reason']) || '').trim();
    var status = String(col_(lvs, r, 'Status') || '').trim();
    var m;
    leaveDateRe_.lastIndex = 0;
    while ((m = leaveDateRe_.exec(datesStr))) {
      var mo = MONTH_IDX_[m[1].slice(0,3).toLowerCase()]; if (mo == null) continue;
      var day = parseInt(m[2], 10);
      var yr = m[3] ? parseInt(m[3], 10) : nowYear_;
      var d = new Date(yr, mo, day); if (isNaN(d.getTime())) continue;
      var iso = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var frac = (m[4]||'').toLowerCase().indexOf('half') >= 0 ? 0.5 : 1;
      leaveRows.push([iso, i, leaveType || 'Leave', frac, purpose, status || 'Unknown']);
    }
  });

  // team daily aggregates for the AHT benchmark (anonymised)
  var tprod = {}, tq = {};
  util.forEach(function(r){ tprod[r[0]] = (tprod[r[0]] || 0) + r[3]; });           // production hours
  queue.forEach(function(r){ if (!tq[r[0]]) tq[r[0]] = [0,0,0]; tq[r[0]][0] += r[2]+r[3]+r[4]; tq[r[0]][1] += r[3]; tq[r[0]][2] += r[4]; });
  var teamUtilDaily = Object.keys(tprod).map(function(d){ return [d, r4_(tprod[d])]; });               // [date, prodHours]
  var teamQueueDaily = Object.keys(tq).map(function(d){ return [d, tq[d][0], tq[d][1], tq[d][2]]; });    // [date, taskCount, trafficking, liveQC]

  var viewer = canon(getViewerEmail());
  var isLead = LEAD_EMAILS.map(function(x){ return normEmail_(x); }).indexOf(viewer) >= 0;
  // Access is restricted to people listed on the Team List tab (Col B = Email).
  var isTeamMember = !!ID.teamEmails[viewer];
  if (!isLead && !isTeamMember) {
    return {
      noAccess: true, viewer: viewer, isLead: false, canExport: false, unknown: true,
      capacity: CAPACITY_HOURS, weights: KPI_WEIGHTS, teamSize: 0,
      teamUtilDaily: [], teamQueueDaily: [],
      emails: [], names: [], countries: [], queue: [], rej: [], util: [], ext: [], score: [],
      chk: [], pktByIdx: [], pktRows: [], selfdevRows: [], compRows: [], leaveRows: [],
      generated: Utilities.formatDate(new Date(), tz, 'd MMM yyyy, HH:mm') + ' (' + tz + ')'
    };
  }
  // Export is restricted to the named lead only
  var canExport = normEmail_(viewer) === normEmail_(EXPORT_LEAD_EMAIL);

  var payload = {
    capacity: CAPACITY_HOURS, viewer: viewer, isLead: isLead, canExport: canExport, ahtBasis: AHT_BASIS,
    teamUtilDaily: teamUtilDaily, teamQueueDaily: teamQueueDaily, teamSize: emails.length,
    generated: Utilities.formatDate(new Date(), tz, 'd MMM yyyy, HH:mm') + ' (' + tz + ')'
  };
  payload.weights = KPI_WEIGHTS;
  if (isLead) {
    payload.emails = emails; payload.names = names; payload.countries = countries;
    payload.queue = queue; payload.rej = rej; payload.util = util; payload.ext = ext; payload.score = score;
    payload.chk = chk; payload.pktByIdx = pktByIdx; payload.pktRows = pktRows;
    payload.selfdevRows = selfdevRows; payload.compRows = compRows; payload.paRows = paRows;
    payload.trainByIdx = trainByIdx; payload.compByIdx = compByIdx; payload.leaveRows = leaveRows; payload.unknown = false;
  } else {
    var vi = idx.hasOwnProperty(viewer) ? idx[viewer] : -1;
    function mine(a){ return a.filter(function(r){ return r[1] === vi; }).map(function(r){ var c = r.slice(); c[1] = 0; return c; }); }
    payload.emails = vi >= 0 ? [emails[vi]] : []; payload.names = vi >= 0 ? [names[vi]] : [];
    payload.countries = vi >= 0 ? [countries[vi]] : [];
    payload.queue = vi >= 0 ? mine(queue) : []; payload.rej = vi >= 0 ? mine(rej) : [];
    payload.util = vi >= 0 ? mine(util) : []; payload.ext = vi >= 0 ? mine(ext) : []; payload.score = vi >= 0 ? mine(score) : [];
    payload.chk = vi >= 0 ? mine(chk) : []; payload.pktByIdx = vi >= 0 ? [pktByIdx[vi]] : [];
    payload.pktRows = vi >= 0 ? mine(pktRows) : [];
    payload.selfdevRows = vi >= 0 ? mine(selfdevRows) : [];
    payload.compRows = vi >= 0 ? mine(compRows) : [];
    payload.paRows = vi >= 0 ? mine(paRows) : [];
    payload.trainByIdx = vi >= 0 ? [trainByIdx[vi]] : [];
    payload.compByIdx = vi >= 0 ? [compByIdx[vi]] : [];
    payload.leaveRows = vi >= 0 ? mine(leaveRows) : [];
    payload.unknown = vi < 0;
  }
  return payload;
}

/* ============================================================
 *  EXPORT — Corporate Scorecard to a new Google Sheet
 *  Restricted to EXPORT_LEAD_EMAIL. Mirrors the Monthly
 *  Performance Review layout (one row per agent, grouped KPIs).
 *  Called from the dashboard "Export to Google Sheet" button.
 * ============================================================ */
const WORKFLOW_LABEL = 'PCoE EMEA';
function bandPctX_(p){ return bandPct_SOT(p); }
function bandUtilX_(p){ return bandUtil_SOT(p); }
function weekKeyX_(iso){ var d=new Date(iso+'T12:00:00Z'); var off=(d.getUTCDay()+6)%7; d.setUTCDate(d.getUTCDate()-off); return d.toISOString().slice(0,10); }
function inFilter_(iso, p){
  if(!iso) return false;
  if(!p) return true;
  if(p.quarters && p.quarters.length && p.quarters.indexOf(qKeyOf_(iso))<0) return false;
  if(p.months && p.months.length && p.months.indexOf(iso.slice(0,7))<0) return false;
  if(p.weeks && p.weeks.length && p.weeks.indexOf(weekKeyX_(iso))<0) return false;
  if(p.from && iso<p.from) return false;
  if(p.to && iso>p.to) return false;
  return true;
}
function periodLabel_(p){
  if(!p) return 'All periods';
  if(p.months && p.months.length) return p.months.map(function(m){ var a=m.split('-'); var MO=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; return MO[+a[1]-1]+' '+a[0]; }).join(', ');
  if(p.quarters && p.quarters.length) return p.quarters.join(', ');
  if(p.weeks && p.weeks.length) return p.weeks.length+' week(s)';
  if(p.from || p.to) return (p.from||'…')+' to '+(p.to||'…');
  return 'All periods';
}
function pctStr_(v){ return (Math.round(v*100)/100).toFixed(2)+'%'; }

function exportScorecardSheet(period) {
  var viewer = (function(){ try { return normEmail_(getViewerEmail()); } catch(e){ return ''; } })();
  if (viewer !== normEmail_(EXPORT_LEAD_EMAIL))
    throw new Error('Export is restricted to ' + EXPORT_LEAD_EMAIL + '.');

  var data = getData();
  var p = period || null;
  var pass = function(iso){ return inFilter_(iso, p); };

  var W = { launch:.15, qa:.15, checklist:.10, pkt:.10, util:.15, pa:.15, training:.10, compliance:.10 };
  var rows = [];

  var wantCountry = (p && p.country) ? String(p.country).trim().toLowerCase() : '';
  data.emails.forEach(function(email, i){
    // country filter from the dashboard dropdown ('' = all countries)
    if (wantCountry && String((data.countries && data.countries[i]) || '').trim().toLowerCase() !== wantCountry) return;
    // queue aggregates
    var assigned=0, traf=0, liveQC=0, rej=0, cil=0;
    data.queue.forEach(function(r){ if(r[1]===i && pass(r[0])){ assigned+=r[2]; traf+=r[3]; liveQC+=r[4]; rej+=r[5]; cil+=r[6]; } });
    var task = assigned + traf + liveQC;
    // external errors + critical misses (both from the External Errors / Issue Type column)
    var extC=0, critC=0;
    data.ext.forEach(function(r){ if(r[1]!==i || !pass(r[0])) return; if(r[2]==='External') extC++; else if(r[2]==='Critical') critC++; });
    // utilization
    var prodH=0, totH=0, dayset={};
    data.util.forEach(function(r){ if(r[1]===i && pass(r[0])){ prodH+=r[3]; totH+=r[4]; dayset[r[0]]=1; } });
    var days = Object.keys(dayset).length;
    // checklist misses
    var misses=0; (data.chk||[]).forEach(function(r){ if(r[1]===i && (r[0]==='' || pass(r[0]))) misses++; });
    // process-adherence CILs (from the Process Adherence detail sheet)
    var cilPA=0; (data.paRows||[]).forEach(function(r){ if(r[1]===i && (r[0]==='' || pass(r[0]))) cilPA++; });
    // pkt (all-time avg per person)
    var pkt = (data.pktByIdx && data.pktByIdx[i] != null) ? data.pktByIdx[i] : null;
    // training (Self Development) — on-time completion %
    var sdDone=0, sdOnT=0;
    (data.selfdevRows||[]).forEach(function(r){ if(r[1]===i && (r[0]==='' || pass(r[0])) && r[4]!=null){ sdDone++; if(r[4]<=0) sdOnT++; } });
    var trainPct = sdDone ? (sdOnT/sdDone)*100 : 100;
    var trainR = (data.trainByIdx && data.trainByIdx[i] != null) ? data.trainByIdx[i] : 5;
    // compliance — count of entries (fewer is better); rating from compBand_
    var cmTot=0;
    (data.compRows||[]).forEach(function(r){ if(r[1]===i && (r[0]==='' || pass(r[0]))) cmTot++; });
    var compR = (data.compByIdx && data.compByIdx[i] != null) ? data.compByIdx[i] : compBand_(cmTot);
    var compPct = (compR/5)*100;
    // latest P / PG rating within filter
    var pr='', pgr='';
    data.score.forEach(function(r){ if(r[1]===i && pass(r[0])){ pr=r[2]; pgr=r[3]; } });

    // skip people with no activity at all in this period
    if (task===0 && totH===0 && extC===0 && critC===0 && misses===0 && sdDone===0 && cmTot===0) return;

    // Critical misses count toward Internal QA (alongside internal rejections), not
    // Launch Accuracy — Launch Accuracy is External-errors-only in this export.
    var rejForQA  = rej + critC;
    var launchPct = traf ? (1-extC/traf)*100 : 100;
    var iqPct     = traf ? (1-rejForQA/traf)*100 : 100;
    var chkEntered= task - misses;
    var chkPct    = task ? (chkEntered/task)*100 : 100;
    var paPct     = task ? Math.max(0,(1-cilPA/task))*100 : 100;
    var utilPct   = days ? (totH/(CAPACITY_HOURS*days))*100 : 0;
    var pktPct    = pkt==null ? 100 : pkt;

    var launchR = extC>=1 ? 1 : 5;
    var iqR     = traf ? bandPctX_(iqPct) : 5;
    var chkR    = misses>=2 ? 1 : misses===1 ? 2 : 5;
    var pktR    = pkt==null ? 5 : bandPctX_(pktPct);
    var utilR   = days ? bandUtilX_(utilPct) : 5;
    var paR     = bandPctX_(paPct);

    var overall = launchR*W.launch + iqR*W.qa + chkR*W.checklist + pktR*W.pkt +
                  utilR*W.util + paR*W.pa + trainR*W.training + compR*W.compliance;
    var overallScore = (overall/5)*100;
    var contrib = function(w, ach){ return pctStr_(w * Math.min(ach,100)); };

    rows.push([
      WORKFLOW_LABEL, data.names[i], pctStr_(overallScore), Math.round(overall*100)/100,
      // Launch: Rejected, Total Tasks, Achieved, Target, %, Rating
      extC, traf, pctStr_(launchPct), '100%', contrib(W.launch,launchPct), launchR,
      // Internal QA: Rejected (internal rejections + critical misses), Total Tasks, Achieved, Target, %, Rating
      rejForQA, traf, pctStr_(iqPct), '98%', contrib(W.qa,iqPct), iqR,
      // Checklist: Task Count, Checklist Entered, %, # of CILs, Rating
      task, chkEntered, pctStr_(chkPct), cilPA, chkR,
      // PKT: Achieved, Target, %, Rating
      pctStr_(pktPct), '100%', contrib(W.pkt,pktPct), pktR,
      // Utilization: Achieved, Target, %, Rating
      pctStr_(utilPct), '85%', contrib(W.util,utilPct), utilR,
      // Process Adherence: Process Adherence, Target, %, Rating
      pctStr_(paPct), '100%', contrib(W.pa,paPct), paR,
      // Training: Achieved, Target, %, Rating
      pctStr_(trainPct), '100%', contrib(W.training,trainPct), trainR,
      // Compliance: Achieved, Target, %, Rating
      pctStr_(compPct), '100%', contrib(W.compliance,compPct), compR,
      // P / PG / QPR
      pr, pgr, Math.round(overall*10)/10,
      // Critical misses in this period (already folded into the Internal QA Rejected count above)
      critC
    ]);
  });

  // sort by Agent name
  rows.sort(function(a,b){ return String(a[1]).localeCompare(String(b[1])); });

  var label = periodLabel_(p);
  if (p && p.country) label += ' · ' + p.country;
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var nss = SpreadsheetApp.create('Scorecard Export — ' + label + ' — ' + stamp);
  var sh = nss.getActiveSheet(); sh.setName('Scorecards');
  var NC = 45;
  function blank(){ var a=[]; for(var j=0;j<NC;j++) a.push(''); return a; }

  // Header rows (mirror the Monthly Performance Review layout)
  var h1=blank(); h1[0]=label; h1[4]='Core Delivery Excellence'; h1[25]='Client & Business Impact'; h1[33]='Self Development'; h1[37]='Compliance'; h1[41]='Rating'; h1[42]='PG Rating';
  var h2=blank(); h2[0]='Workflow'; h2[1]='Agent name'; h2[2]='Overall Score'; h2[3]='Rating';
  h2[4]='Launch Accuracy'; h2[10]='Internal QA Score'; h2[16]='Checklist Adherence'; h2[21]='PKT';
  h2[25]='Utilization'; h2[29]='Process Adherence - Overall'; h2[33]='Training/Certification Completion'; h2[37]='Compliance'; h2[43]='Overall Rating QPR';
  h2[44]='Critical Misses';
  var h3=blank(); h3[4]='15%'; h3[10]='15%'; h3[16]='10%'; h3[21]='10%'; h3[25]='15%'; h3[29]='15%'; h3[33]='10%'; h3[37]='10%';
  var h4=blank();
  var sub=['Rejected','Total Tasks','Achieved','Target','%','Rating']; sub.forEach(function(s,j){ h4[4+j]=s; h4[10+j]=s; });
  ['Task Count','Checklist Entered','%','# of CILs','Rating'].forEach(function(s,j){ h4[16+j]=s; });
  ['Achieved','Target','%','Rating'].forEach(function(s,j){ h4[21+j]=s; h4[25+j]=s; h4[33+j]=s; h4[37+j]=s; });
  ['Process Adherence','Target','%','Rating'].forEach(function(s,j){ h4[29+j]=s; });

  var all = [h1,h2,h3,h4].concat(rows.length ? rows : [blank()]);
  sh.getRange(1,1,all.length,NC).setValues(all);

  // formatting
  sh.getRange(1,1,4,NC).setFontWeight('bold').setBackground('#0c6170').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(2,1,1,2).setHorizontalAlignment('left');
  sh.setFrozenRows(4); sh.setFrozenColumns(2);
  if (rows.length) sh.getRange(5,1,rows.length,NC).setHorizontalAlignment('center');
  sh.getRange(5,2,Math.max(rows.length,1),1).setHorizontalAlignment('left');
  sh.autoResizeColumns(1, NC);

  return nss.getUrl();
}

/* ============================================================
 *  EMAIL SYSTEM — daily · weekly · monthly · quarterly
 *  Four distinct, professional templates with LIVE insights
 *  computed from each person's real numbers.
 *
 *  TEST FIRST (emails only the lead — no one else is touched):
 *     testDaily()   testWeekly()   testMonthly()   testQuarterly()
 *  When happy, wire time-driven triggers to the trigger* fns:
 *     Weekly   -> triggerWeeklyEmail   (Mondays)
 *     Monthly  -> triggerMonthlyEmail  (before the 3rd)
 *     Quarterly-> triggerQuarterlyEmail(before the 5th)
 * ============================================================ */

// EMEA Hub teal palette — shared header gradient, subtle per-cadence accent
var EMAIL_HEADER = 'linear-gradient(100deg,#0b4d5c 0%,#11808c 52%,#34b3ad 100%)';
var EMAIL_BG = '#eef4f8';
var EMAIL_THEME = {
  daily:     { accent:'#0e8f93', accent2:'#0c6170', tag:'Daily Pulse',                cadence:'Daily',     emoji:'⚡' },
  weekly:    { accent:'#11808c', accent2:'#0b4d5c', tag:'Weekly Scorecard',           cadence:'Weekly',    emoji:'📊' },
  monthly:   { accent:'#0c6170', accent2:'#083d49', tag:'Monthly Performance Review',  cadence:'Monthly',   emoji:'📈' },
  quarterly: { accent:'#0a5566', accent2:'#06323d', tag:'Quarterly Business Review',   cadence:'Quarterly', emoji:'🏆' }
};

/* convenience test entry points — send ONLY to the lead running the script */
function testDaily()     { sendScorecardEmail('daily', true); }
function testWeekly()    { sendScorecardEmail('weekly', true); }
function testMonthly()   { sendScorecardEmail('monthly', true); }
function testQuarterly() { sendScorecardEmail('quarterly', true); }
function testEmails()    { sendScorecardEmail('daily', true); } // legacy alias

/* ---------- reporting window for a timeframe (relative to now) ---------- */
function windowFor_(timeframe, tz) {
  var now = new Date();
  function iso(d){ return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  var start, end = now, label;
  if (timeframe === 'daily') {
    start = new Date(now); start.setDate(now.getDate() - 1);
    label = iso(start);
  } else if (timeframe === 'weekly') {
    start = new Date(now); start.setDate(now.getDate() - 7);
    label = iso(start) + '  to  ' + iso(end);
  } else if (timeframe === 'monthly') {
    start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    end   = new Date(now.getFullYear(), now.getMonth(), 0);
    label = Utilities.formatDate(start, tz, 'MMMM yyyy');
  } else { // quarterly -> previous quarter
    var q = Math.floor(now.getMonth() / 3), ps = (q - 1) * 3, y = now.getFullYear();
    if (ps < 0) { ps = 9; y--; }
    start = new Date(y, ps, 1); end = new Date(y, ps + 3, 0);
    label = 'Q' + (Math.floor(ps / 3) + 1) + ' ' + y;
  }
  return { start: iso(start), end: iso(end), label: label };
}

/* ---------- per-user metrics over a window ---------- */
function metricsForUser_(data, vi, win) {
  function inWin(d){ return !win || (d >= win.start && d <= win.end); }
  function inScope(ri){ return vi === null || ri === vi; }
  var q  = data.queue.filter(function(r){ return inScope(r[1]) && inWin(r[0]); });
  var u  = data.util .filter(function(r){ return inScope(r[1]) && inWin(r[0]); });
  var x  = data.ext  .filter(function(r){ return inScope(r[1]) && inWin(r[0]); });
  var rj = data.rej  .filter(function(r){ return inScope(r[1]) && inWin(r[0]); });
  var assigned=0, traf=0, liveQC=0, rej=0, cil=0;
  q.forEach(function(r){ assigned+=r[2]; traf+=r[3]; liveQC+=r[4]; rej+=r[5]; cil+=r[6]; });
  var task = assigned + traf + liveQC;
  var extC = x.filter(function(r){ return r[2] === 'External'; }).length;
  var crit = x.filter(function(r){ return r[2] === 'Critical'; }).length;
  var totH=0, prodH=0, nonH=0, pd={};
  u.forEach(function(r){ nonH+=r[2]; prodH+=r[3]; totH+=r[4]; pd[r[0]+'|'+r[1]]=1; });
  var personDays = Object.keys(pd).length;
  return {
    task:task, assigned:assigned, traf:traf, liveQC:liveQC, rej:rej, cil:cil,
    extC:extC, crit:crit,
    iq: traf ? (1 - rej / traf) * 100 : null,
    eq: traf ? (1 - extC / traf) * 100 : null,
    totH:totH, prodH:prodH, nonH:nonH, personDays:personDays,
    util: personDays ? (totH / (CAPACITY_HOURS * personDays)) * 100 : 0,
    ahtSec: task ? prodH * 3600 / task : 0,
    hasData: (q.length + u.length + x.length) > 0
  };
}

/* ---------- corporate scorecard: parameter ratings + overall ---------- */
function teamAHTsec_(data, win) {
  function inWin(d){ return !win || (d >= win.start && d <= win.end); }
  var prod = 0, denom = 0;
  data.teamUtilDaily.forEach(function(r){ if (inWin(r[0])) prod += r[1]; });
  data.teamQueueDaily.forEach(function(r){ if (inWin(r[0])) denom += (AHT_BASIS === 'taskcount' ? r[1] : AHT_BASIS === 'liveqc' ? r[3] : r[2]); });
  return denom ? prod * 3600 / denom : 0;
}
var RATING_LABELS = { 5:'Top Quartile', 4:'Exceeds', 3:'Meets', 2:'Below', 1:'Bottom Quartile' };
var KPI_META_ = [
  { k:'launch',     label:'Launch Accuracy',          bucket:'Core Delivery Excellence' },
  { k:'qa',         label:'Internal QA Score',         bucket:'Core Delivery Excellence' },
  { k:'checklist',  label:'Checklist Adherence',       bucket:'Core Delivery Excellence' },
  { k:'pkt',        label:'PKT',                       bucket:'Core Delivery Excellence' },
  { k:'pa',         label:'Process Adherence',         bucket:'Client & Business Impact' },
  { k:'util',       label:'Utilization',               bucket:'Client & Business Impact' },
  { k:'training',   label:'Training / Certification',  bucket:'Self Development' },
  { k:'compliance', label:'Compliance',                bucket:'Compliance' }
];
function qKeyOf_(iso){ var y=iso.slice(0,4), mo=+iso.slice(5,7); return y+'-Q'+(Math.floor((mo-1)/3)+1); }
function prevQ_(qk){ var p=qk.split('-Q'), y=+p[0], q=+p[1]-1; if(q<1){q=4;y--;} return y+'-Q'+q; }
function qLabel_(qk){ var p=qk.split('-Q'); return 'Q'+p[1]+' '+p[0]; }
/* ===== single source of truth for the band cut-offs (guardrail) ===== */
function bandPct_SOT(p){ return p>=100?5:p>=99?4:p>=98?3:p>=97?2:1; }            // QA · PKT · Process Adherence
function bandUtil_SOT(p){ return p>=85?5:p>=80?4:p>=75?3:p>=70?2:1; }            // Utilization
function compBand_(count){ count=count||0; return count<=0?5:count===1?4:count===2?3:count===3?2:1; } // Compliance (count of entries)
/* Launch Accuracy (External Errors, Issue Type column):
 *   eC  = External-type count in the CURRENT quarter
 *   eP  = External-type count 1 quarter back · eP2 = 2 quarters back
 *   cC  = Critical-type count in the CURRENT quarter (only counts when the current quarter itself is clean of External)
 * Base band = how far back the last External error was (current / last / 2-back / clean).
 * A clean current quarter is then pulled down by Critical misses logged this quarter — more
 * criticals in a "should be higher" band drag it further toward 1. */
function launchRating_(eC, eP, eP2, cC) {
  cC = cC || 0;
  if (eC >= 1) return 1;                                            // external error this quarter — always bottom
  if (eP >= 1) return cC > 1 ? 1 : cC === 1 ? 2 : 3;                 // external last quarter only
  if (eP2 >= 1) return cC > 2 ? 1 : cC === 2 ? 2 : cC === 1 ? 3 : 4; // external 2 quarters back only
  return cC > 3 ? 1 : cC === 3 ? 2 : cC === 2 ? 3 : cC === 1 ? 4 : 5; // clean for 2+ quarters
}
function trainBandDelay_(delay){ return delay==null?1:delay<=0?5:delay===1?4:delay===2?3:delay===3?2:1; } // Training (days late)

/* Run from the editor: asserts every band returns the intended rating. Fails loud on drift. */
function validateScorecard() {
  var fails = [], ok = 0;
  function expect(label, got, want){ if (got!==want) fails.push(label+': got '+got+', expected '+want); else ok++; }
  // Compliance — count based
  expect('compliance 0', compBand_(0), 5);
  expect('compliance 1', compBand_(1), 4);
  expect('compliance 2', compBand_(2), 3);
  expect('compliance 3', compBand_(3), 2);
  expect('compliance 4', compBand_(4), 1);
  expect('compliance 9', compBand_(9), 1);
  // Launch Accuracy (External Errors + Critical downgrade)
  expect('launch cur ext', launchRating_(1,0,0,0), 1);
  expect('launch cur ext+crit', launchRating_(2,0,0,3), 1);
  expect('launch clean 3+', launchRating_(0,0,0,0), 5);
  expect('launch clean 3+ crit1', launchRating_(0,0,0,1), 4);
  expect('launch clean 3+ crit2', launchRating_(0,0,0,2), 3);
  expect('launch clean 3+ crit3', launchRating_(0,0,0,3), 2);
  expect('launch clean 3+ crit4', launchRating_(0,0,0,4), 1);
  expect('launch 2-back only', launchRating_(0,0,1,0), 4);
  expect('launch 2-back crit1', launchRating_(0,0,1,1), 3);
  expect('launch 2-back crit2', launchRating_(0,0,1,2), 2);
  expect('launch 2-back crit3', launchRating_(0,0,1,3), 1);
  expect('launch last only', launchRating_(0,1,0,0), 3);
  expect('launch last crit1', launchRating_(0,1,0,1), 2);
  expect('launch last crit2', launchRating_(0,1,0,2), 1);
  // Utilization
  expect('util 85', bandUtil_SOT(85), 5); expect('util 84.99', bandUtil_SOT(84.99), 4);
  expect('util 80', bandUtil_SOT(80), 4); expect('util 75', bandUtil_SOT(75), 3);
  expect('util 70', bandUtil_SOT(70), 2); expect('util 69.9', bandUtil_SOT(69.9), 1);
  // % based KPIs
  expect('pct 100', bandPct_SOT(100), 5); expect('pct 99', bandPct_SOT(99), 4);
  expect('pct 98', bandPct_SOT(98), 3); expect('pct 97', bandPct_SOT(97), 2); expect('pct 96.9', bandPct_SOT(96.9), 1);
  // Training — days late
  expect('train ontime', trainBandDelay_(0), 5); expect('train 1d', trainBandDelay_(1), 4);
  expect('train 2d', trainBandDelay_(2), 3); expect('train 3d', trainBandDelay_(3), 2);
  expect('train 4d', trainBandDelay_(4), 1); expect('train pending', trainBandDelay_(null), 1);
  var msg = fails.length ? ('❌ '+fails.length+' FAILED:\n'+fails.join('\n')) : ('✅ All '+ok+' scorecard band checks passed.');
  Logger.log(msg); return msg;
}

/* training / compliance rating from per-person arrays (vi===null => team average) */
function fromByIdx_(arr, vi) {
  if (!arr) return 5;
  if (vi === null) { var v = arr.filter(function(x){ return x != null; }); return v.length ? Math.round(v.reduce(function(a,b){return a+b;},0)/v.length) : 5; }
  return arr[vi] != null ? arr[vi] : 5;
}
function trainRating_(data, vi){ return fromByIdx_(data.trainByIdx, vi); }
function compRating_(data, vi){ return fromByIdx_(data.compByIdx, vi); }

/* Quarterly ratings — Launch & Checklist use prior-quarter history.
   Current quarter = latest quarter inside the report window, else latest overall. */
function computeRatings_(data, vi, m, win) {
  function bandPct(p){ return p>=100?5:p>=99?4:p>=98?3:p>=97?2:1; }
  var inScopeR = function(ri){ return vi === null || ri === vi; };
  function bucket(rows, ok, val){ var o={}; rows.forEach(function(r){ if(!inScopeR(r[1])||!ok(r))return; var k=qKeyOf_(r[0]); o[k]=(o[k]||0)+val(r); }); return o; }
  var hasDate = function(r){ return r[0] && r[0]!==''; };
  var extQ  = bucket(data.ext,  function(r){ return hasDate(r) && r[2]==='External'; }, function(){return 1;});
  var critQ = bucket(data.ext,  function(r){ return hasDate(r) && r[2]==='Critical'; }, function(){return 1;});
  var missQ = bucket(data.chk||[], hasDate, function(){return 1;});
  var cilQ  = bucket(data.paRows||[], hasDate, function(){return 1;});  // CILs from Process Adherence sheet (detail-accurate)
  var taskQ = bucket(data.queue, hasDate, function(r){return r[2]+r[3]+r[4];});
  var rejQ  = bucket(data.queue, hasDate, function(r){return r[5];});
  var trafQ = bucket(data.queue, hasDate, function(r){return r[3];});
  var prodQ = bucket(data.util,  hasDate, function(r){return r[3];});
  var totQ  = bucket(data.util,  hasDate, function(r){return r[4];});  // total logged hours (utilization basis)

  // current quarter
  var qs = {};
  [data.queue, data.ext, data.util].forEach(function(a){ a.forEach(function(r){ if(inScopeR(r[1]) && hasDate(r) && (!win || (r[0]>=win.start && r[0]<=win.end))) qs[qKeyOf_(r[0])]=1; }); });
  var inWinQs = Object.keys(qs).sort();
  var allQs = {};
  [data.queue, data.ext, data.util].forEach(function(a){ a.forEach(function(r){ if(inScopeR(r[1]) && hasDate(r)) allQs[qKeyOf_(r[0])]=1; }); });
  var allList = Object.keys(allQs).sort();
  var cur = inWinQs.length ? inWinQs[inWinQs.length-1] : (allList.length ? allList[allList.length-1] : null);

  var pkt;
  if (vi === null) { var pv = (data.pktByIdx||[]).filter(function(x){ return x != null; }); pkt = pv.length ? pv.reduce(function(a,b){return a+b;},0)/pv.length : null; }
  else pkt = (data.pktByIdx && data.pktByIdx[vi] != null) ? data.pktByIdx[vi] : null;
  var sc, paPct=100;
  if (!cur) {
    sc = { launch:5, qa:5, checklist:5, pkt:5, pa:5, util:5, training:5, compliance:5 };
  } else {
    var p1 = prevQ_(cur), p2 = prevQ_(p1), G = function(o,k){ return o[k]||0; };
    var eC=G(extQ,cur), eP=G(extQ,p1), eP2=G(extQ,p2);
    var cC=G(critQ,cur);
    var mC=G(missQ,cur), mP=G(missQ,p1), mP2=G(missQ,p2);
    var taskC=G(taskQ,cur), cilC=G(cilQ,cur);
    paPct = taskC ? Math.max(0,(1-cilC/taskC))*100 : 100;
    var utilDaysQ_ = {}, seenUD_ = {};
    data.util.forEach(function(r){ if(inScopeR(r[1]) && hasDate(r)){ var k=qKeyOf_(r[0]), dk=r[0]+'|'+r[1]; if(!seenUD_[dk]){ seenUD_[dk]=1; utilDaysQ_[k]=(utilDaysQ_[k]||0)+1; } } });
    var utilDayCnt = utilDaysQ_[cur] || 0;
    var utilTotH_ = G(totQ, cur);   // total logged hours (matches the dashboard Utilization card)
    var utilPct = utilDayCnt ? (utilTotH_ / (CAPACITY_HOURS * utilDayCnt)) * 100 : 0;
    var rejC=G(rejQ,cur), trafC=G(trafQ,cur), iqPct = trafC ? (1-rejC/trafC)*100 : null;
    sc = {
      launch:    launchRating_(eC, eP, eP2, cC),
      qa:        iqPct==null ? 5 : bandPct(iqPct),
      checklist: mC>=2 ? 1 : mC===1 ? 2 : (mP===0&&mP2===0 ? 5 : mP===0 ? 4 : 3),
      pkt:       pkt==null ? 5 : bandPct(pkt),
      pa:        bandPct(paPct),
      util:      !utilDayCnt ? 5 : (utilPct>=85?5:utilPct>=80?4:utilPct>=75?3:utilPct>=70?2:1),
      training:  trainRating_(data, vi),
      compliance: compRating_(data, vi)
    };
  }
  var overall = 0; for (var k in KPI_WEIGHTS) overall += sc[k] * KPI_WEIGHTS[k];
  return { sc: sc, overall: overall, paPct: paPct, pkt: pkt, cur: cur };
}
function overallLabel_(o){ return o>=4.5?'Exceptional':o>=3.5?'Exceeds':o>=2.5?'Meets':o>=1.5?'Below Expectations':'Needs Improvement'; }

/* ---------- live insight + suggestion engine (shared with dashboard logic) ---------- */
function buildInsights_(m) {
  var ins = [], sug = [];
  if (m.iq != null) {
    if (m.iq >= 99)      ins.push({ tone:'good',  text:'Internal quality is excellent at ' + m.iq.toFixed(1) + '% — work is clearing QC cleanly.' });
    else if (m.iq >= 97) ins.push({ tone:'watch', text:'Internal quality is ' + m.iq.toFixed(1) + '%, just shy of the 99% top band.' });
    else { ins.push({ tone:'risk', text:'Internal quality is ' + m.iq.toFixed(1) + '% (' + m.rej + ' rejection' + (m.rej===1?'':'s') + ' on ' + m.traf + ' items).' });
           sug.push('Pair with QC on the top rejection reason this period to lift internal quality back above 99%.'); }
  }
  if (m.extC > 0) { ins.push({ tone:'risk', text:m.extC + ' external, client-facing error' + (m.extC>1?'s':'') + ' logged — launch accuracy is the priority.' });
                    sug.push('Review each external error and add a checklist guard so the same root cause cannot recur.'); }
  else if (m.traf > 0) ins.push({ tone:'good', text:'Zero external errors across ' + m.traf + ' trafficked launch' + (m.traf===1?'':'es') + ' — accuracy at 100%.' });
  if (m.personDays > 0) {
    if (m.util < 85)      { ins.push({ tone:'watch', text:'Utilization is ' + m.util.toFixed(0) + '%, below the 85% healthy band.' });
                            sug.push('Log all production time in the tracker and flag any blockers pulling you off task.'); }
    else if (m.util > 115)  ins.push({ tone:'watch', text:'Utilization is ' + m.util.toFixed(0) + '% — high; protect a sustainable pace.' });
    else                    ins.push({ tone:'good',  text:'Utilization is healthy at ' + m.util.toFixed(0) + '%.' });
  }
  if (m.cil > 0) { ins.push({ tone:'watch', text:m.cil + ' CIL' + (m.cil>1?'s':'') + ' recorded — currently captured under Process Adherence.' });
                   sug.push('Close out open CILs and confirm the corrective action is documented for the rating cycle.'); }
  if (!ins.length) ins.push({ tone:'good', text:'No activity in this window yet — metrics will populate as the trackers fill in.' });
  if (!sug.length) sug.push('Keep the streak going and document value-adds to strengthen the next P/PG rating.');
  return { items: ins, suggestions: sug };
}

/* ---------- small formatters ---------- */
function fmtPct_(v){ return v == null ? '—' : v.toFixed(1) + '%'; }
function fmtAht_(sec){ sec = Math.max(0, Math.round(sec)); var h = Math.floor(sec/3600), m = Math.floor((sec%3600)/60), s = sec%60; return (h>0?h+':'+(m<10?'0':'')+m:m) + ':' + (s<10?'0':'') + s; }
function escHtml_(v){ return String(v == null ? '' : v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function fmtDate_(iso){ if (!iso) return '—'; var d = new Date(iso + 'T12:00:00Z'); var MO=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; return d.getUTCDate() + ' ' + MO[d.getUTCMonth()] + ' ' + d.getUTCFullYear(); }

/* ---------- email HTML builder ---------- */
function buildScorecardHTML_(name, timeframe, m, R, win) {
  var t = EMAIL_THEME[timeframe];
  var ins = buildInsights_(m);
  var deep = (timeframe === 'monthly' || timeframe === 'quarterly'); // show ratings + compliance block

  function tile(label, value, note) {
    return '<td width="50%" style="padding:6px;" valign="top">' +
      '<div style="background:#f5f9fb;border:1px solid #e3edf2;border-radius:12px;padding:16px 18px;">' +
      '<div style="font-size:10px;color:#7d96a0;text-transform:uppercase;letter-spacing:1px;">' + label + '</div>' +
      '<div style="font-size:26px;font-weight:600;color:' + t.accent + ';margin-top:4px;font-family:Georgia,serif;">' + value + '</div>' +
      (note ? '<div style="font-size:11px;color:#9a9a9d;margin-top:3px;">' + note + '</div>' : '') +
      '</div></td>';
  }
  var tiles =
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin:18px 0;"><tr>' +
      tile('Task Count', m.task, m.assigned + ' assigned · ' + m.traf + ' traffic · ' + m.liveQC + ' QC') +
      tile('Internal Quality', fmtPct_(m.iq), m.rej + ' rejection' + (m.rej===1?'':'s')) +
    '</tr><tr>' +
      tile('External Quality', fmtPct_(m.eq), m.extC + ' external error' + (m.extC===1?'':'s')) +
      tile('Utilization', m.personDays ? m.util.toFixed(0) + '%' : '—', m.totH.toFixed(1) + ' hrs logged') +
    '</tr><tr>' +
      tile('Production', m.prodH.toFixed(1) + ' hrs', 'of ' + m.totH.toFixed(1) + ' hrs logged') +
      tile('CIL Count', m.cil, 'under Process Adherence') +
    '</tr></table>';

  var dashUrl = dashboardUrl_();
  var ctaHTML = dashUrl ? (
    '<div style="text-align:center;margin-top:24px;">' +
      '<a href="' + dashUrl + '" target="_blank" style="display:inline-block;background:' + EMAIL_HEADER + ';color:#ffffff;text-decoration:none;font-size:13px;font-weight:700;letter-spacing:.3px;padding:13px 28px;border-radius:30px;box-shadow:0 6px 16px rgba(12,97,112,.28);">Open the full dashboard →</a>' +
    '</div>') : '';
  var toneColor = { good:'#0e8f93', watch:'#c0892f', risk:'#d9534f' };
  var insightHTML = ins.items.map(function(i){
    return '<tr><td width="8" valign="top" style="padding-top:6px;"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + (toneColor[i.tone]||'#7a7a7d') + ';"></span></td>' +
           '<td style="padding:4px 0 4px 10px;font-size:13px;color:#3f3f42;line-height:1.5;">' + i.text + '</td></tr>';
  }).join('');
  var sugHTML = ins.suggestions.map(function(s){ return '<li style="margin-bottom:6px;">' + s + '</li>'; }).join('');

  // Overall rating banner (all timeframes)
  var ratingColor = function(s){ return s>=4 ? '#0e8f93' : s===3 ? '#c0892f' : '#d9534f'; };
  var overallBanner =
    '<div style="margin:18px 0;background:linear-gradient(135deg,#0c6170,#0b4d5c);border-radius:14px;padding:20px 22px;color:#fff;">' +
      '<div style="font-size:10px;color:rgba(255,255,255,.65);text-transform:uppercase;letter-spacing:1.5px;">Overall Rating · per Corporate Scorecard' + (R.cur ? ' · ' + qLabel_(R.cur) : '') + '</div>' +
      '<div style="margin:4px 0 2px;"><span style="font-size:40px;font-weight:600;color:#5fd6cf;font-family:Georgia,serif;">' + R.overall.toFixed(2) + '</span>' +
      '<span style="font-size:16px;color:rgba(255,255,255,.5);"> / 5.00</span>' +
      '<span style="float:right;margin-top:14px;font-size:13px;font-weight:600;">' + overallLabel_(R.overall) + '</span></div>' +
    '</div>';

  // Full parameter breakdown (monthly + quarterly)
  var rows = KPI_META_.map(function(o){
    var s = R.sc[o.k], col = ratingColor(s);
    return '<tr><td style="padding:9px 6px;font-size:13px;color:#3f3f42;border-bottom:1px solid #f2f2f2;">' + o.label + '</td>' +
      '<td style="padding:9px 6px;text-align:right;border-bottom:1px solid #f2f2f2;"><span style="background:' + col + '18;color:' + col + ';font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;">' + s + '/5 · ' + RATING_LABELS[s] + '</span></td>' +
      '<td style="padding:9px 6px;text-align:right;font-size:12px;color:#9a9a9d;border-bottom:1px solid #f2f2f2;">' + Math.round((KPI_WEIGHTS[o.k]||0)*100) + '%</td></tr>';
  }).join('');
  var paramTable = deep ?
    ('<h3 style="border-bottom:1px solid #eef0ef;padding-bottom:10px;font-weight:500;color:#3f3f42;margin:26px 0 6px;font-family:Georgia,serif;">Parameter ratings</h3>' +
     '<table width="100%" cellpadding="0" cellspacing="0"><thead><tr>' +
       '<th style="text-align:left;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Parameter</th>' +
       '<th style="text-align:right;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Rating</th>' +
       '<th style="text-align:right;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Weight</th>' +
     '</tr></thead><tbody>' + rows + '</tbody></table>') : '';
  var ratingBlock = overallBanner + paramTable;

  return '' +
  '<div style="background:' + EMAIL_BG + ';padding:24px 12px;">' +
  '<div style="font-family:\'Segoe UI\',Tahoma,Verdana,sans-serif;max-width:640px;margin:0 auto;background:#ffffff;border-radius:18px;overflow:hidden;box-shadow:0 12px 40px rgba(12,97,112,.14);color:#15323a;">' +
    '<div style="background:' + EMAIL_HEADER + ';padding:32px 30px 28px;text-align:center;color:#fff;">' +
      '<div style="font-weight:700;font-size:18px;letter-spacing:-.2px;">EMEA<span style="color:#bff3ef;">Social</span></div>' +
      '<h1 style="margin:14px 0 4px;font-size:24px;font-weight:500;font-family:Georgia,serif;">Operations Scorecard</h1>' +
      '<div style="display:inline-block;margin-top:8px;background:rgba(255,255,255,.18);color:#ffffff;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;padding:5px 14px;border-radius:20px;">' + t.emoji + ' ' + t.tag + '</div>' +
      '<div style="margin-top:12px;font-size:12px;color:rgba(255,255,255,.72);">Reporting period · ' + win.label + '</div>' +
    '</div>' +
    '<div style="padding:28px 30px;">' +
      '<p style="font-size:16px;margin:0 0 4px;">Hi <strong>' + name + '</strong>,</p>' +
      '<p style="color:#7d96a0;line-height:1.6;margin:0;">Here is your ' + t.cadence.toLowerCase() + ' performance snapshot, measured against the corporate KPI framework.</p>' +
      tiles +
      ratingBlock +
      '<h3 style="border-bottom:1px solid #e3edf2;padding-bottom:10px;font-weight:500;color:#3a5660;margin:28px 0 12px;font-family:Georgia,serif;">' + t.emoji + ' Insights</h3>' +
      '<table width="100%" cellpadding="0" cellspacing="0">' + insightHTML + '</table>' +
      '<div style="margin-top:22px;background:rgba(20,145,155,.08);border-left:4px solid ' + t.accent + ';padding:16px 18px;border-radius:0 10px 10px 0;">' +
        '<div style="font-size:12px;font-weight:700;color:' + t.accent2 + ';text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">💡 Suggested actions</div>' +
        '<ul style="margin:0;padding-left:18px;font-size:13px;color:#3a5660;line-height:1.5;">' + sugHTML + '</ul>' +
      '</div>' +
      ctaHTML +
    '</div>' +
    '<div style="background:#f5f9fb;padding:18px;text-align:center;font-size:11px;color:#7d96a0;border-top:1px solid #e3edf2;">' +
      '<p style="margin:0;">Automated ' + t.cadence + ' report · EMEA Social Posting Hub — Operations KPI Dashboard</p>' +
      '<p style="margin:5px 0 0;">© 2026 EMEA Social. All rights reserved.</p>' +
    '</div>' +
  '</div>' +
  '</div>';
}

/* ---------- send ----------
 * TEST mode: always sends exactly ONE preview email to whoever runs the script
 *   (the lead). If the lead has their own data we use it; otherwise we preview
 *   using the first team member's numbers — but the email still goes to the lead.
 * LIVE mode: one email per person, to their own address. */
function sendScorecardEmail(timeframe, testMode) {
  var data = getData();
  var ss = ss_(); var tz = ss.getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var win = windowFor_(timeframe, tz);
  var sent = [];
  // Nigeria team members get their manager + SME cc'd on every individual scorecard — LIVE sends only,
  // so a TEST preview never touches real managers.
  var nigeriaEmails = {};
  teamListEmailsByCountry_(NIGERIA_COUNTRY_VALUE).forEach(function(e){ nigeriaEmails[e] = true; });

  function buildAndSend(vi, toEmail) {
    var name = data.names[vi];
    var useWin = win;
    var m = metricsForUser_(data, vi, win);
    var windowLabel = win.label;
    if (!m.hasData) {                            // graceful fallback so test emails still show real numbers
      useWin = null;
      m = metricsForUser_(data, vi, null);
      windowLabel = win.label + ' · showing all available data';
    }
    var R = computeRatings_(data, vi, m, useWin);
    var theme = EMAIL_THEME[timeframe];
    var subject = (testMode ? '[TEST] ' : '') + theme.emoji + ' ' + theme.cadence + ' Ops KPI Scorecard — ' + name;
    var htmlBody = buildScorecardHTML_(name, timeframe, m, R, { start: win.start, end: win.end, label: windowLabel });
    var mail = { to: toEmail, subject: subject, htmlBody: htmlBody, name: EMAIL_SENDER_NAME };
    if (!testMode && nigeriaEmails[normEmail_(toEmail)]) mail.cc = NIGERIA_REPORT_TO.join(',');
    MailApp.sendEmail(mail);
    sent.push(name + ' <' + toEmail + '>' + (mail.cc ? ' (cc Nigeria manager/SME)' : ''));
  }

  // Team-overall scorecard for the lead(s) — whole-team numbers + dashboard link.
  function buildAndSendLead(toEmail) {
    var useWin = win;
    var m = metricsForUser_(data, null, win);
    var windowLabel = win.label;
    if (!m.hasData) { useWin = null; m = metricsForUser_(data, null, null); windowLabel = win.label + ' · showing all available data'; }
    var R = computeRatings_(data, null, m, useWin);
    var theme = EMAIL_THEME[timeframe];
    var subject = (testMode ? '[TEST] ' : '') + theme.emoji + ' ' + theme.cadence + ' Ops KPI Scorecard — Team Overall';
    var htmlBody = buildScorecardHTML_('Team Lead', timeframe, m, R, { start: win.start, end: win.end, label: windowLabel });
    MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: htmlBody, name: EMAIL_SENDER_NAME });
    sent.push('Team Overall <' + toEmail + '>');
  }

  if (testMode) {
    var viewer = getViewerEmail();
    if (!viewer) return 'TEST aborted: could not read your email (Session.getActiveUser). Authorize the script and retry.';
    if (!data.emails.length) return 'TEST aborted: no people found in the data tabs.';
    var vi = data.emails.indexOf(viewer);          // lead's own row if present…
    if (vi < 0) vi = 0;                            // …otherwise preview with the first person
    buildAndSend(vi, viewer);                       // …individual preview…
    buildAndSendLead(viewer);                       // …plus the team-overall lead email
  } else {
    data.emails.forEach(function(email, vi) { buildAndSend(vi, email); });
    LEAD_EMAILS.forEach(function(leadEmail){ if (leadEmail) buildAndSendLead(normEmail_(leadEmail)); }); // team overall to each lead
  }

  logEmailStatus_(timeframe, testMode, win.label, sent);
  return 'Sent ' + sent.length + ' ' + timeframe + ' email(s) to: ' + sent.join(', ') + (testMode ? '  (TEST — delivered to you only)' : '') + '.';
}

/* ---------- status logging (Email Status tab) ---------- */
function logEmailStatus_(timeframe, testMode, periodLabel, recipients) {
  try {
    var ss = ss_();
    var sh = ss.getSheetByName(TABS.emailStatus);
    if (!sh) { sh = ss.insertSheet(TABS.emailStatus); }
    if (sh.getLastRow() === 0)
      sh.appendRow(['Timestamp', 'Timeframe', 'Period', 'Mode', 'Status', 'Recipients', 'Detail']);
    sh.appendRow([
      new Date(), timeframe, periodLabel,
      testMode ? 'TEST' : 'LIVE', 'Sent',
      recipients.length, recipients.join('; ')
    ]);
  } catch (e) {}
}

function triggerDailyEmail() { sendScorecardEmail('daily', false); }
function triggerWeeklyEmail() { sendScorecardEmail('weekly', false); }
function triggerMonthlyEmail() { sendScorecardEmail('monthly', false); }
function triggerQuarterlyEmail() { sendScorecardEmail('quarterly', false); }

/* ============================================================
 *  TEAM DIGEST EMAIL — one email to the whole team (not per-person)
 *  Cadences: daily (09:00 IST) · weekly (Monday) · monthly (2nd).
 *  Anonymized: Internal Rejections and External Errors (incl. Critical
 *  misses) are never attributed to a person — only category/context,
 *  so the team stays aware without singling anyone out.
 *
 *  APPROVAL WORKFLOW: sendTeamDigestEmail(cadence, true) ALWAYS
 *  previews to DIGEST_PREVIEW_EMAIL only — nobody else is touched.
 *  Run testTeamDigestDaily() / Weekly() / Monthly() from the editor,
 *  review the template, then wire the trigger* functions below to
 *  time-driven triggers via installTeamDigestTriggers() (run once).
 *
 *  DEDUPE: every LIVE send is logged to the Email Status tab; a
 *  second attempt for the same cadence+period is skipped so a
 *  re-fired trigger can never double-send.
 * ============================================================ */
const DISCREPANCY_FORM_URL = 'https://script.google.com/a/macros/mediamint.com/s/AKfycbzMc4VyDgodJjV18LBxa5YTsShtAGkdYEYELLjbUAWQocWElEdBkRUkiEsB3YFBlUI4/exec';
const DIGEST_PREVIEW_EMAIL = 'sairam.konda@mediamint.com';
const DIGEST_THEME = { accent: '#0e8f93', accent2: '#0c6170' };

function testTeamDigestDaily()     { return sendTeamDigestEmail('daily', true); }
function testTeamDigestWeekly()    { return sendTeamDigestEmail('weekly', true); }
function testTeamDigestMonthly()   { return sendTeamDigestEmail('monthly', true); }
function testTeamDigestQuarterly() { return sendTeamDigestEmail('quarterly', true); }

function triggerTeamDigestDaily() { sendTeamDigestEmail('daily', false); }
function triggerTeamDigestWeekly() {
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  if (Utilities.formatDate(new Date(), tz, 'EEEE') !== 'Monday') return; // guard, in case the trigger fires off-schedule
  sendTeamDigestEmail('weekly', false);
}
function triggerTeamDigestMonthly() {
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  if (Utilities.formatDate(new Date(), tz, 'd') !== '2') return; // only actually sends on the 2nd
  sendTeamDigestEmail('monthly', false);
}
function triggerTeamDigestQuarterly() {
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var d = Utilities.formatDate(new Date(), tz, 'd'), mo = +Utilities.formatDate(new Date(), tz, 'M');
  if (d !== '2' || [1,4,7,10].indexOf(mo) < 0) return; // only fires on the 2nd of Jan/Apr/Jul/Oct
  sendTeamDigestEmail('quarterly', false);
}

/* Run ONCE from the editor to install the daily/weekly/monthly/quarterly triggers.
 * IMPORTANT: time-driven trigger timing follows the Apps Script PROJECT
 * time zone (⚙ Project Settings ▸ Time zone in the editor) — set that to
 * Asia/Kolkata first, or these fire at 9:00 in the wrong zone. */
function installTeamDigestTriggers() {
  ['triggerTeamDigestDaily', 'triggerTeamDigestWeekly', 'triggerTeamDigestMonthly', 'triggerTeamDigestQuarterly'].forEach(function(fn){
    ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
  });
  ScriptApp.newTrigger('triggerTeamDigestDaily').timeBased().everyDays(1).atHour(9).nearMinute(0).create();
  ScriptApp.newTrigger('triggerTeamDigestWeekly').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).nearMinute(0).create();
  ScriptApp.newTrigger('triggerTeamDigestMonthly').timeBased().everyDays(1).atHour(9).nearMinute(0).create();   // self-guards to the 2nd only
  ScriptApp.newTrigger('triggerTeamDigestQuarterly').timeBased().everyDays(1).atHour(9).nearMinute(0).create(); // self-guards to the 2nd of Jan/Apr/Jul/Oct
  return 'Installed. Daily @~9am, Weekly (Mon) @~9am, Monthly (guarded to the 2nd) @~9am, Quarterly (guarded to the 2nd of Jan/Apr/Jul/Oct) @~9am — confirm the project time zone is Asia/Kolkata first.';
}

/* ---------- reporting window + prior comparison window for a cadence ---------- */
function digestWindow_(cadence, tz) {
  var now = new Date();
  function iso(d){ return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  if (cadence === 'daily') {
    var y = new Date(now); y.setDate(now.getDate() - 1);
    var py = new Date(now); py.setDate(now.getDate() - 2);
    return { start: iso(y), end: iso(y), label: 'Yesterday · ' + iso(y),
             cmpStart: iso(py), cmpEnd: iso(py), cmpLabel: 'day before (' + iso(py) + ')' };
  }
  if (cadence === 'weekly') {
    var end = new Date(now); end.setDate(now.getDate() - 1);        // yesterday (Sunday, when the trigger runs Monday)
    var start = new Date(end); start.setDate(end.getDate() - 6);
    var cEnd = new Date(start); cEnd.setDate(start.getDate() - 1);
    var cStart = new Date(cEnd); cStart.setDate(cEnd.getDate() - 6);
    return { start: iso(start), end: iso(end), label: iso(start) + ' to ' + iso(end),
             cmpStart: iso(cStart), cmpEnd: iso(cEnd), cmpLabel: 'prior week (' + iso(cStart) + ' to ' + iso(cEnd) + ')' };
  }
  if (cadence === 'quarterly') {
    var q = Math.floor(now.getMonth() / 3), ps = (q - 1) * 3, y = now.getFullYear();
    if (ps < 0) { ps = 9; y--; }
    var qs2 = new Date(y, ps, 1), qe2 = new Date(y, ps + 3, 0);
    var cps = ps - 3, cy = y;
    if (cps < 0) { cps = 9; cy--; }
    var cqs = new Date(cy, cps, 1), cqe = new Date(cy, cps + 3, 0);
    return { start: iso(qs2), end: iso(qe2), label: 'Q' + (Math.floor(ps / 3) + 1) + ' ' + y,
             cmpStart: iso(cqs), cmpEnd: iso(cqe), cmpLabel: 'Q' + (Math.floor(cps / 3) + 1) + ' ' + cy };
  }
  // monthly (default/fallback)
  var s = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var e = new Date(now.getFullYear(), now.getMonth(), 0);
  var cs = new Date(now.getFullYear(), now.getMonth() - 2, 1);
  var ce = new Date(now.getFullYear(), now.getMonth() - 1, 0);
  return { start: iso(s), end: iso(e), label: Utilities.formatDate(s, tz, 'MMMM yyyy'),
           cmpStart: iso(cs), cmpEnd: iso(ce), cmpLabel: Utilities.formatDate(cs, tz, 'MMMM yyyy') };
}

/* short, subject-line-friendly period label: "Aug 19" (daily) · "Aug 13 – Aug 19" (weekly) · "August 2026" (monthly) */
function digestSubjectPeriod_(cadence, win, tz) {
  function fmt(iso, pattern){ return Utilities.formatDate(new Date(iso + 'T12:00:00Z'), tz, pattern); }
  if (cadence === 'daily') return fmt(win.end, 'MMM d');
  if (cadence === 'weekly') return fmt(win.start, 'MMM d') + ' – ' + fmt(win.end, 'MMM d');
  return win.label; // monthly is already "MMMM yyyy"
}

/* canonical, de-duplicated list of emails straight from the Team List tab (Col: Email) —
 * the digest's recipient list is drawn from here, not from who happened to have activity. */
function teamListEmails_() {
  var ts = readSheet_(TABS.team);
  var seen = {}, out = [];
  ts.rows.forEach(function(r){
    var e = normEmail_(col_(ts, r, 'Email'));
    if (e && !seen[e]) { seen[e] = 1; out.push(e); }
  });
  return out;
}

/* ---------- region / country tagging (Team List, Col C = "Country") ----------
 * Powers the Nigeria-dedicated report: which Team List members are in which
 * country, so team-wide numbers can be scoped down to just that country. */
function teamCountryMap_() {
  var ts = readSheet_(TABS.team);
  var out = {};
  ts.rows.forEach(function(r){
    var e = normEmail_(col_(ts, r, 'Email')); if (!e) return;
    var c = String(col_(ts, r, 'Country', ['Region']) || '').trim();
    if (c) out[e] = c;
  });
  return out;
}
/* idx set (into data.emails/data.names) of everyone tagged with the given country */
function regionIdxSet_(data, countryValue) {
  var cmap = teamCountryMap_();
  var want = String(countryValue || '').trim().toLowerCase();
  var set = {};
  data.emails.forEach(function(e, i){
    var c = cmap[normEmail_(e)];
    if (c && c.toLowerCase() === want) set[i] = true;
  });
  return set;
}
/* de-duplicated Team List emails for a given country (for recipient lists, dedupe, etc.) */
function teamListEmailsByCountry_(countryValue) {
  var ts = readSheet_(TABS.team);
  var want = String(countryValue || '').trim().toLowerCase();
  var seen = {}, out = [];
  ts.rows.forEach(function(r){
    var e = normEmail_(col_(ts, r, 'Email')); if (!e || seen[e]) return;
    var c = String(col_(ts, r, 'Country', ['Region']) || '').trim().toLowerCase();
    if (c === want) { seen[e] = 1; out.push(e); }
  });
  return out;
}
/* Returns a shallow clone of getData()'s payload scoped to one country: every
 * per-row array (queue/util/ext/rej/chk/paRows/selfdevRows/compRows/pktRows/
 * score) is filtered down to rows whose idx is in that country, and the
 * per-person arrays (pktByIdx/trainByIdx/compByIdx) are masked to null for
 * everyone outside it. emails/names stay full-length so indices still line
 * up — agents outside the region simply end up with zero activity and are
 * skipped automatically by teamAgentRows_() etc. */
function filterDataToRegion_(data, idxSet) {
  function inSet(i){ return !!idxSet[i]; }
  function filterRows(arr){ return (arr || []).filter(function(r){ return inSet(r[1]); }); }
  function maskByIdx(arr){ return (arr || []).map(function(v, i){ return inSet(i) ? v : null; }); }
  var out = {};
  for (var k in data) out[k] = data[k];
  out.queue = filterRows(data.queue);
  out.rej = filterRows(data.rej);
  out.util = filterRows(data.util);
  out.ext = filterRows(data.ext);
  out.score = filterRows(data.score);
  out.chk = filterRows(data.chk);
  out.paRows = filterRows(data.paRows);
  out.selfdevRows = filterRows(data.selfdevRows);
  out.compRows = filterRows(data.compRows);
  out.pktRows = filterRows(data.pktRows);
  out.leaveRows = filterRows(data.leaveRows);
  out.pktByIdx = maskByIdx(data.pktByIdx);
  out.trainByIdx = maskByIdx(data.trainByIdx);
  out.compByIdx = maskByIdx(data.compByIdx);
  return out;
}

/* ---------- whole-team (all agents) aggregate metrics over a window ---------- */
function teamWindowMetrics_(data, start, end) {
  function inW(d){ return d >= start && d <= end; }
  var q = data.queue.filter(function(r){ return inW(r[0]); });
  var u = data.util.filter(function(r){ return inW(r[0]); });
  var x = data.ext.filter(function(r){ return inW(r[0]); });
  var rj = data.rej.filter(function(r){ return inW(r[0]); });
  var assigned=0, traf=0, liveQC=0, rej=0, cil=0;
  q.forEach(function(r){ assigned+=r[2]; traf+=r[3]; liveQC+=r[4]; rej+=r[5]; cil+=r[6]; });
  var task = assigned + traf + liveQC;
  var extC = x.filter(function(r){ return r[2]==='External'; }).length;
  var critC = x.filter(function(r){ return r[2]==='Critical'; }).length;
  var totH=0, prodH=0, nonH=0, pd={};
  u.forEach(function(r){ nonH+=r[2]; prodH+=r[3]; totH+=r[4]; pd[r[0]+'|'+r[1]]=1; });
  var personDays = Object.keys(pd).length;
  var rejByCat = {};
  rj.forEach(function(r){ var c=r[2]||'Other'; rejByCat[c]=(rejByCat[c]||0)+r[3]; });
  var extByContext = {}; // anonymized context: WHAT happened and why, never who it belongs to (never Task Name)
  x.forEach(function(r){
    var issue = String(r[5]||'').trim() || 'No details logged';
    var k = (r[2]||'Other') + ' — ' + issue;
    extByContext[k] = (extByContext[k]||0) + 1;
  });
  var pv = (data.pktByIdx||[]).filter(function(v){ return v != null; });
  var pkt = pv.length ? pv.reduce(function(a,b){return a+b;},0) / pv.length : null;
  var denom = CAPACITY_HOURS * personDays;
  return {
    task:task, assigned:assigned, traf:traf, liveQC:liveQC, rej:rej, cil:cil, extC:extC, critC:critC,
    totH:totH, prodH:prodH, nonH:nonH, personDays:personDays,
    iq: traf ? (1-rej/traf)*100 : null, eq: traf ? (1-extC/traf)*100 : null,
    util: personDays ? (totH/denom)*100 : 0,
    prodUtil: personDays ? (prodH/denom)*100 : 0,
    nonProdUtil: personDays ? (nonH/denom)*100 : 0,
    ahtSec: task ? prodH*3600/task : 0,
    pkt: pkt, rejByCat: rejByCat, extByContext: extByContext
  };
}

/* ---------- per-agent contribution rows over a window (Individual Contributions table) ---------- */
function teamAgentRows_(data, start, end) {
  function inW(d){ return d >= start && d <= end; }
  var rows = data.emails.map(function(email, i){
    var q = data.queue.filter(function(r){ return r[1]===i && inW(r[0]); });
    var u = data.util.filter(function(r){ return r[1]===i && inW(r[0]); });
    var x = data.ext.filter(function(r){ return r[1]===i && inW(r[0]); });
    var assigned=0, traf=0, liveQC=0, rej=0;
    q.forEach(function(r){ assigned+=r[2]; traf+=r[3]; liveQC+=r[4]; rej+=r[5]; });
    var task = assigned + traf + liveQC;
    var extC = x.filter(function(r){ return r[2]==='External'; }).length;
    var critC = x.filter(function(r){ return r[2]==='Critical'; }).length;
    var totH=0, prodH=0, nonH=0, pd={};
    u.forEach(function(r){ nonH+=r[2]; prodH+=r[3]; totH+=r[4]; pd[r[0]]=1; });
    var days = Object.keys(pd).length;
    var denom = CAPACITY_HOURS * days;
    return {
      name: data.names[i], task:task, traf:traf, rej:rej, extC:extC, critC:critC,
      iq: traf ? (1-rej/traf)*100 : null,
      util: days ? (totH/denom)*100 : null,
      prodUtil: days ? (prodH/denom)*100 : null,
      nonProdUtil: days ? (nonH/denom)*100 : null,
      ahtSec: task ? prodH*3600/task : 0
    };
  }).filter(function(r){ return r.task>0 || r.rej>0 || r.extC>0 || r.critC>0 || r.util!=null; });
  rows.sort(function(a,b){ return a.name.localeCompare(b.name); });
  return rows;
}

/* ---------- top-N helper (categories / context, by count desc) ---------- */
function topN_(map, n) {
  return Object.keys(map).sort(function(a,b){ return map[b]-map[a]; }).slice(0, n);
}
function digestAlreadySent_(key, periodLabel) {
  try {
    var sh = ss_().getSheetByName(TABS.emailStatus);
    if (!sh) return false;
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++)
      if (String(vals[i][1])===key && String(vals[i][2])===periodLabel && String(vals[i][3])==='LIVE' && String(vals[i][4])==='Sent') return true;
  } catch (e) {}
  return false;
}

/* ---------- per-incident / per-record DETAIL rows for a window (full detail, names included) ----------
 * Replaces the old anonymized trend-tables with straight per-record listings so the
 * digest shows exactly who/when/what for every rejection, external/critical error,
 * compliance entry, process-adherence miss, checklist miss and leave in the window. */
function teamExtDetailRows_(data, start, end) {
  function inW(d){ return d >= start && d <= end; }
  return data.ext.filter(function(r){ return inW(r[0]); }).map(function(r){
    return { date:r[0], name:data.names[r[1]]||'—', type:r[2]||'—', task:r[3]||'', link:r[4]||'', issue:r[5]||'' };
  }).sort(function(a,b){ return a.date<b.date?1:a.date>b.date?-1:0; });
}
function teamRejDetailRows_(data, start, end) {
  function inW(d){ return d >= start && d <= end; }
  return data.rej.filter(function(r){ return inW(r[0]); }).map(function(r){
    return { date:r[0], name:data.names[r[1]]||'—', category:r[2]||'Other', count:r[3]||0 };
  }).sort(function(a,b){ return a.date<b.date?1:a.date>b.date?-1:0; });
}
function teamComplianceDetailRows_(data, start, end) {
  function inW(d){ return d && d>=start && d<=end; }
  return (data.compRows||[]).filter(function(r){ return inW(r[0]); }).map(function(r){
    return { date:r[0], name:data.names[r[1]]||'—', type:r[2]||'Compliance', details:r[3]||'' };
  }).sort(function(a,b){ return a.date<b.date?1:a.date>b.date?-1:0; });
}
function teamPADetailRows_(data, start, end) {
  function inW(d){ return d && d>=start && d<=end; }
  return (data.paRows||[]).filter(function(r){ return inW(r[0]); }).map(function(r){
    return { date:r[0], name:data.names[r[1]]||'—', missedBy:r[2]||'Other', columns:r[3]||'—', link:r[4]||'' };
  }).sort(function(a,b){ return a.date<b.date?1:a.date>b.date?-1:0; });
}
function teamChecklistDetailRows_(data, start, end) {
  function inW(d){ return d && d>=start && d<=end; }
  return (data.chk||[]).filter(function(r){ return inW(r[0]); }).map(function(r){
    return { date:r[0], name:data.names[r[1]]||'—', category:r[2]||'Uncategorised' };
  }).sort(function(a,b){ return a.date<b.date?1:a.date>b.date?-1:0; });
}
function teamLeaveDetailRows_(data, start, end) {
  function inW(d){ return d && d>=start && d<=end; }
  return (data.leaveRows||[]).filter(function(r){ return inW(r[0]); }).map(function(r){
    return { date:r[0], name:data.names[r[1]]||'—', type:r[2]||'Leave', fraction:r[3]||1, purpose:r[4]||'', status:r[5]||'Unknown' };
  }).sort(function(a,b){ return a.date<b.date?1:a.date>b.date?-1:0; });
}

/* ---------- team-level insights + suggestions (anonymized) ---------- */
function buildTeamInsights_(m, cmp) {
  var ins=[], sug=[];
  if (m.iq != null) {
    var d = cmp.iq != null ? m.iq - cmp.iq : 0;
    ins.push({ tone: m.iq>=99?'good':m.iq>=97?'watch':'risk',
      text: 'Team Internal Quality is ' + m.iq.toFixed(1) + '%' + (cmp.iq!=null ? ' (' + (d>=0?'+':'') + d.toFixed(1) + ' pt vs ' + cmp.__label + ')' : '') +
            ' across ' + m.rej + ' rejection' + (m.rej===1?'':'s') + ' on ' + m.traf + ' trafficked items.' });
    if (m.iq < 99) sug.push('Review the top 1–2 rejection categories below with the team in the next huddle — a quick root-cause pass on the biggest category usually moves this the most.');
  }
  var extD = m.extC - (cmp.extC||0), critD = m.critC - (cmp.critC||0);
  if (m.extC > 0 || m.critC > 0) {
    ins.push({ tone:'risk', text: m.extC + ' external error' + (m.extC===1?'':'s') + ' and ' + m.critC + ' critical miss' + (m.critC===1?'':'es') + ' logged' +
      (' (' + (extD>=0?'+':'') + extD + ' external, ' + (critD>=0?'+':'') + critD + ' critical vs prior period)') + ' — launch accuracy is the priority.' });
    sug.push('Walk through each external/critical item as a team and add a checklist guard so the same root cause cannot recur.');
  } else {
    ins.push({ tone:'good', text: 'Zero external errors and zero critical misses this period across ' + m.traf + ' trafficked launch' + (m.traf===1?'':'es') + '.' });
  }
  if (m.personDays > 0) {
    if (m.util < 85) { ins.push({ tone:'watch', text:'Team utilization is ' + m.util.toFixed(0) + '%, below the 85% healthy band.' });
                        sug.push('Confirm production time is fully logged and flag any recurring blockers pulling the team off task.'); }
    else ins.push({ tone:'good', text:'Team utilization is healthy at ' + m.util.toFixed(0) + '%.' });
  }
  if (!sug.length) sug.push('Keep the streak going — call out the categories driving improvement in the next stand-up.');
  return { items: ins, suggestions: sug };
}

/* ---------- discrepancy CTA ---------- */
function discrepancyButtonHTML_(sentAtLabel) {
  return '<div style="margin-top:22px;text-align:center;">' +
    '<a href="' + DISCREPANCY_FORM_URL + '" target="_blank" style="display:inline-block;background:#d9534f;color:#ffffff;text-decoration:none;font-size:13px;font-weight:700;letter-spacing:.3px;padding:12px 26px;border-radius:30px;box-shadow:0 6px 16px rgba(217,83,79,.28);">🚩 Spot a discrepancy? Raise it here</a>' +
    '<div style="font-size:11px;color:#9a9a9d;margin-top:8px;">SLA: 24 hours from this report — sent ' + sentAtLabel + '</div>' +
  '</div>';
}

/* ---------- single metric tile (handles nullable current/comparison values) ---------- */
function digestTile_(accent, label, curNum, prevNum, higherIsBetter, fmt) {
  fmt = fmt || function(v){ return v==null ? '—' : String(Math.round(v*100)/100); };
  var arrowHtml = '';
  if (curNum != null && prevNum != null) {
    var diff = curNum - prevNum;
    var good = diff===0 ? null : (diff>0) === higherIsBetter;
    var col = diff===0 ? '#9a9a9d' : (good ? '#0e8f93' : '#d9534f');
    var tri = diff===0 ? '→' : (diff>0 ? '▲' : '▼');
    arrowHtml = ' <span style="color:' + col + ';font-size:11px;font-weight:700;">' + tri + (diff===0?'':' '+fmt(Math.abs(diff))) + '</span>';
  }
  return '<td width="33%" style="padding:6px;" valign="top">' +
    '<div style="background:#f5f9fb;border:1px solid #e3edf2;border-radius:12px;padding:14px 16px;">' +
    '<div style="font-size:10px;color:#7d96a0;text-transform:uppercase;letter-spacing:1px;">' + label + '</div>' +
    '<div style="font-size:22px;font-weight:700;color:' + accent + ';margin-top:4px;font-family:Georgia,serif;">' + fmt(curNum) + arrowHtml + '</div>' +
    (prevNum != null ? '<div style="font-size:11px;color:#9a9a9d;margin-top:3px;">vs ' + fmt(prevNum) + '</div>' : '') +
    '</div></td>';
}

/* ---------- shared per-record detail table (Date / Person / ... — full detail, names included) ----------
 * cols: [{label, right, wrap, get(row)->html}]. Empty rows render an empty-state paragraph instead. */
function digestDetailTable_(title, note, cols, rows, emptyMsg) {
  var heading = '<h3 style="border-bottom:1px solid #eef0ef;padding-bottom:8px;font-weight:500;color:#3f3f42;margin:24px 0 6px;font-family:Georgia,serif;">' +
    title + (note ? ' <span style="font-size:11px;color:#9a9a9d;font-weight:400;">' + note + '</span>' : '') + '</h3>';
  if (!rows.length) return heading + '<p style="font-size:12.5px;color:#7d96a0;margin-top:0;">' + emptyMsg + '</p>';
  var head = cols.map(function(c){
    return '<th style="text-align:' + (c.right?'right':'left') + ';font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;' + (c.wrap?'':'white-space:nowrap;') + '">' + c.label + '</th>';
  }).join('');
  var body = rows.map(function(r){
    return '<tr>' + cols.map(function(c){
      return '<td style="padding:7px 6px;font-size:12px;color:#3f3f42;border-bottom:1px solid #f2f2f2;vertical-align:top;' + (c.right?'text-align:right;':'') + (c.wrap?'':'white-space:nowrap;') + '">' + c.get(r) + '</td>';
    }).join('') + '</tr>';
  }).join('');
  return heading + '<table width="100%" cellpadding="0" cellspacing="0"><thead><tr>' + head + '</tr></thead><tbody>' + body + '</tbody></table>';
}

/* ---------- digest HTML builder ---------- */
function buildDigestHTML_(cadence, win, m, cmp, insSug, includeRatings, R, sentAtLabel, agentRows, data, regionLabel) {
  cmp.__label = win.cmpLabel;
  var cadeLabel = cadence.charAt(0).toUpperCase() + cadence.slice(1);
  var teamTitle = (regionLabel ? regionLabel + ' Team' : 'Team') + ' Operations Digest';
  var teamIntro = regionLabel
    ? 'Here is the ' + regionLabel + ' team’s ' + cadence + ' snapshot, in full detail.'
    : 'Here is the whole-team ' + cadence + ' snapshot, in full detail.';
  var accent = DIGEST_THEME.accent, accent2 = DIGEST_THEME.accent2;
  var pctFmt = function(v){ return v==null ? '—' : v.toFixed(1) + '%'; };
  var intFmt = function(v){ return v==null ? '—' : String(Math.round(v)); };
  var hrFmt  = function(v){ return v==null ? '—' : v.toFixed(1) + ' hrs'; };

  var tiles = '<table width="100%" cellpadding="0" cellspacing="0" style="margin:16px 0;"><tr>' +
    digestTile_(accent, 'Task Count', m.task, cmp.task, true, intFmt) +
    digestTile_(accent, 'Internal Quality', m.iq, cmp.iq, true, pctFmt) +
    digestTile_(accent, 'External Quality', m.eq, cmp.eq, true, pctFmt) +
  '</tr><tr>' +
    digestTile_(accent, 'Internal Rejections', m.rej, cmp.rej, false, intFmt) +
    digestTile_(accent, 'External Errors', m.extC, cmp.extC, false, intFmt) +
    digestTile_(accent, 'Critical Misses', m.critC, cmp.critC, false, intFmt) +
  '</tr><tr>' +
    digestTile_(accent, 'Utilization', m.util, cmp.util, true, pctFmt) +
    digestTile_(accent, 'Production Hrs', m.prodH, cmp.prodH, true, hrFmt) +
    (includeRatings ? digestTile_(accent, 'PKT (team avg)', m.pkt, cmp.pkt, true, pctFmt) : digestTile_(accent, 'CIL Count', m.cil, cmp.cil, false, intFmt)) +
  '</tr><tr>' +
    digestTile_(accent, 'Production Utilization', m.prodUtil, cmp.prodUtil, true, pctFmt) +
    digestTile_(accent, 'Non-Production Utilization', m.nonProdUtil, cmp.nonProdUtil, false, pctFmt) +
    digestTile_(accent, 'Productive AHT', m.ahtSec, cmp.ahtSec, false, function(v){ return v==null ? '—' : fmtAht_(v); }) +
  '</tr></table>';

  // Kept to 6 compact columns (some paired into one cell) so the table fits a ~600px
  // email width without being clipped by the card's overflow:hidden — a 9-column
  // version with every metric in its own cell ran wider than the container.
  var agentRowsHTML = (agentRows||[]).map(function(a){
    return '<tr><td style="padding:6px 5px;font-size:11.5px;font-weight:600;color:#15323a;border-bottom:1px solid #f2f2f2;">' + escHtml_(a.name) + '</td>' +
      '<td style="padding:6px 5px;text-align:right;font-size:11.5px;border-bottom:1px solid #f2f2f2;">' + a.task + '</td>' +
      '<td style="padding:6px 5px;text-align:right;font-size:11.5px;border-bottom:1px solid #f2f2f2;">' + (a.iq==null?'—':a.iq.toFixed(1)+'%') + '</td>' +
      '<td style="padding:6px 5px;text-align:right;font-size:11.5px;border-bottom:1px solid #f2f2f2;">' + a.extC + ' / ' + a.critC + '</td>' +
      '<td style="padding:6px 5px;text-align:right;font-size:11.5px;border-bottom:1px solid #f2f2f2;">' + (a.util==null?'—':a.util.toFixed(1)+'%') + '</td>' +
      '<td style="padding:6px 5px;text-align:right;font-size:11.5px;border-bottom:1px solid #f2f2f2;">' + (a.prodUtil==null?'—':a.prodUtil.toFixed(1)+'%') + ' / ' + (a.nonProdUtil==null?'—':a.nonProdUtil.toFixed(1)+'%') + '</td>' +
      '<td style="padding:6px 5px;text-align:right;font-size:11.5px;border-bottom:1px solid #f2f2f2;">' + fmtAht_(a.ahtSec) + '</td></tr>';
  }).join('');
  var agentTable = agentRowsHTML ? (
    '<h3 style="border-bottom:1px solid #eef0ef;padding-bottom:8px;font-weight:500;color:#3f3f42;margin:24px 0 6px;font-family:Georgia,serif;">Individual Contributions</h3>' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="table-layout:fixed;"><thead><tr>' +
      '<th style="text-align:left;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">Agent</th>' +
      '<th style="text-align:right;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">Tasks</th>' +
      '<th style="text-align:right;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">Int. Qlty</th>' +
      '<th style="text-align:right;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">Ext / Crit</th>' +
      '<th style="text-align:right;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">Util.</th>' +
      '<th style="text-align:right;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">Prod / Non-Prod</th>' +
      '<th style="text-align:right;font-size:9px;color:#9a9a9d;text-transform:uppercase;letter-spacing:.5px;padding:6px 5px;">AHT</th>' +
    '</tr></thead><tbody>' + agentRowsHTML + '</tbody></table>') : '';

  var rejCats = Object.keys(m.rejByCat).sort(function(a,b){ return m.rejByCat[b]-m.rejByCat[a]; });
  var rejRows = rejCats.map(function(cat){
    var c = m.rejByCat[cat], p = cmp.rejByCat[cat]||0, d = c-p;
    return '<tr><td style="padding:7px 6px;font-size:12.5px;color:#3f3f42;border-bottom:1px solid #f2f2f2;">' + escHtml_(cat) + '</td>' +
      '<td style="padding:7px 6px;text-align:right;font-size:12.5px;border-bottom:1px solid #f2f2f2;">' + c + '</td>' +
      '<td style="padding:7px 6px;text-align:right;font-size:11.5px;color:' + (d>0?'#d9534f':d<0?'#0e8f93':'#9a9a9d') + ';border-bottom:1px solid #f2f2f2;">' + (d===0?'→ flat':(d>0?'▲ +':'▼ ')+d) + ' vs prior</td></tr>';
  }).join('');
  var rejTable = rejRows ? (
    '<h3 style="border-bottom:1px solid #eef0ef;padding-bottom:8px;font-weight:500;color:#3f3f42;margin:24px 0 6px;font-family:Georgia,serif;">Internal Rejections — by category</h3>' +
    '<table width="100%" cellpadding="0" cellspacing="0"><thead><tr>' +
      '<th style="text-align:left;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Category</th>' +
      '<th style="text-align:right;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Count</th>' +
      '<th style="text-align:right;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">vs prior period</th>' +
    '</tr></thead><tbody>' + rejRows + '</tbody></table>') :
    '<p style="font-size:12.5px;color:#7d96a0;margin-top:24px;">No internal rejections logged this period.</p>';

  // External Errors & Critical Misses — full per-incident detail, including who it belongs to.
  var extDetailRows = teamExtDetailRows_(data, win.start, win.end);
  var extTable = digestDetailTable_(
    'External Errors &amp; Critical Misses — full detail',
    '(' + extDetailRows.length + ' record' + (extDetailRows.length===1?'':'s') + ')',
    [
      { label:'Date', get:function(r){ return fmtDate_(r.date); } },
      { label:'Person', get:function(r){ return escHtml_(r.name); } },
      { label:'Type', get:function(r){ return escHtml_(r.type); } },
      { label:'Task', wrap:true, get:function(r){ return escHtml_(r.task); } },
      { label:'Link', get:function(r){ return r.link ? '<a href="' + escHtml_(r.link) + '" target="_blank">Open ↗</a>' : '—'; } },
      { label:'What happened', wrap:true, get:function(r){ return escHtml_(r.issue).replace(/\n/g,'<br>'); } }
    ],
    extDetailRows,
    'No external errors or critical misses logged this period.'
  );

  var top3Rej = topN_(m.rejByCat, 3), top3Ext = topN_(m.extByContext, 3);
  var topSummary = (top3Rej.length || top3Ext.length) ? (
    '<div style="margin-top:14px;padding:14px 16px;background:#fbf7ee;border:1px solid #f0e6cc;border-radius:10px;">' +
      '<div style="font-size:11px;font-weight:700;color:#8a6d1f;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;">Top 3 this period</div>' +
      (top3Rej.length ? '<div style="font-size:12px;color:#3a5660;margin-bottom:4px;"><strong>Internal:</strong> ' + top3Rej.map(function(c){ return c+' ('+m.rejByCat[c]+')'; }).join(' · ') + '</div>' : '') +
      (top3Ext.length ? '<div style="font-size:12px;color:#3a5660;">' + '<strong>External / Critical:</strong> ' + top3Ext.map(function(c){ return c+' ('+m.extByContext[c]+')'; }).join(' · ') + '</div>' : '') +
    '</div>') : '';

  var toneColor = { good:'#0e8f93', watch:'#c0892f', risk:'#d9534f' };
  var insightHTML = insSug.items.map(function(i){
    return '<tr><td width="8" valign="top" style="padding-top:6px;"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + (toneColor[i.tone]||'#7a7a7d') + ';"></span></td>' +
           '<td style="padding:4px 0 4px 10px;font-size:13px;color:#3f3f42;line-height:1.5;">' + i.text + '</td></tr>';
  }).join('');
  var sugHTML = insSug.suggestions.map(function(s){ return '<li style="margin-bottom:6px;">' + s + '</li>'; }).join('');

  var ratingColor = function(s){ return s>=4 ? '#0e8f93' : s===3 ? '#c0892f' : '#d9534f'; };
  var ratingBlock = '';
  if (includeRatings && R) {
    var rows = KPI_META_.map(function(o){
      var s = R.sc[o.k], col = ratingColor(s);
      return '<tr><td style="padding:9px 6px;font-size:13px;color:#3f3f42;border-bottom:1px solid #f2f2f2;">' + o.label + '</td>' +
        '<td style="padding:9px 6px;text-align:right;border-bottom:1px solid #f2f2f2;"><span style="background:' + col + '18;color:' + col + ';font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;">' + s + '/5 · ' + RATING_LABELS[s] + '</span></td>' +
        '<td style="padding:9px 6px;text-align:right;font-size:12px;color:#9a9a9d;border-bottom:1px solid #f2f2f2;">' + Math.round((KPI_WEIGHTS[o.k]||0)*100) + '%</td></tr>';
    }).join('');
    ratingBlock =
      '<div style="margin:20px 0;background:linear-gradient(135deg,#0c6170,#0b4d5c);border-radius:14px;padding:20px 22px;color:#fff;">' +
        '<div style="font-size:10px;color:rgba(255,255,255,.65);text-transform:uppercase;letter-spacing:1.5px;">Team Overall Rating · per Corporate Scorecard' + (R.cur ? ' · ' + qLabel_(R.cur) : '') + '</div>' +
        '<div style="margin:4px 0 2px;"><span style="font-size:36px;font-weight:600;color:#5fd6cf;font-family:Georgia,serif;">' + R.overall.toFixed(2) + '</span>' +
        '<span style="font-size:15px;color:rgba(255,255,255,.5);"> / 5.00</span>' +
        '<span style="float:right;margin-top:12px;font-size:13px;font-weight:600;">' + overallLabel_(R.overall) + '</span></div>' +
      '</div>' +
      '<h3 style="border-bottom:1px solid #eef0ef;padding-bottom:10px;font-weight:500;color:#3f3f42;margin:24px 0 6px;font-family:Georgia,serif;">Parameter ratings (includes PKT)</h3>' +
      '<table width="100%" cellpadding="0" cellspacing="0"><thead><tr>' +
        '<th style="text-align:left;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Parameter</th>' +
        '<th style="text-align:right;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Rating</th>' +
        '<th style="text-align:right;font-size:10px;color:#9a9a9d;text-transform:uppercase;letter-spacing:1px;padding:6px;">Weight</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table>';
  }

  // Full per-record detail — internal rejections, compliance, process adherence, checklist, leaves — all with names.
  var rejDetailRows = teamRejDetailRows_(data, win.start, win.end);
  var rejDetailTable = digestDetailTable_(
    'Internal Rejections — full detail', '(' + rejDetailRows.length + ' record' + (rejDetailRows.length===1?'':'s') + ')',
    [
      { label:'Date', get:function(r){ return fmtDate_(r.date); } },
      { label:'Person', get:function(r){ return escHtml_(r.name); } },
      { label:'Category', wrap:true, get:function(r){ return escHtml_(r.category); } },
      { label:'Count', right:true, get:function(r){ return String(r.count); } }
    ],
    rejDetailRows, 'No internal rejections logged this period.'
  );

  var complianceRows = teamComplianceDetailRows_(data, win.start, win.end);
  var complianceTable = digestDetailTable_(
    'Compliance — full detail', '(' + complianceRows.length + ' entr' + (complianceRows.length===1?'y':'ies') + ')',
    [
      { label:'Date', get:function(r){ return fmtDate_(r.date); } },
      { label:'Person', get:function(r){ return escHtml_(r.name); } },
      { label:'Type', get:function(r){ return escHtml_(r.type); } },
      { label:'Details', wrap:true, get:function(r){ return escHtml_(r.details).replace(/\n/g,'<br>'); } }
    ],
    complianceRows, 'No compliance entries logged this period.'
  );

  var paDetailRows = teamPADetailRows_(data, win.start, win.end);
  var paDetailTable = digestDetailTable_(
    'Process Adherence Misses (CILs) — full detail', '(' + paDetailRows.length + ' miss' + (paDetailRows.length===1?'':'es') + ')',
    [
      { label:'Date', get:function(r){ return fmtDate_(r.date); } },
      { label:'Person', get:function(r){ return escHtml_(r.name); } },
      { label:'Missed By', get:function(r){ return escHtml_(r.missedBy); } },
      { label:'Columns Missed', wrap:true, get:function(r){ return escHtml_(r.columns); } },
      { label:'Link', get:function(r){ return r.link ? '<a href="' + escHtml_(r.link) + '" target="_blank">Open ↗</a>' : '—'; } }
    ],
    paDetailRows, 'No process-adherence misses logged this period.'
  );

  var checklistDetailRows = teamChecklistDetailRows_(data, win.start, win.end);
  var checklistDetailTable = digestDetailTable_(
    'Checklist Misses — full detail', '(' + checklistDetailRows.length + ' miss' + (checklistDetailRows.length===1?'':'es') + ')',
    [
      { label:'Date', get:function(r){ return fmtDate_(r.date); } },
      { label:'Person', get:function(r){ return escHtml_(r.name); } },
      { label:'Category', wrap:true, get:function(r){ return escHtml_(r.category); } }
    ],
    checklistDetailRows, 'No checklist misses logged this period.'
  );

  var leaveDetailRows = teamLeaveDetailRows_(data, win.start, win.end);
  var leaveDetailTable = digestDetailTable_(
    'Leaves — full detail', '(' + leaveDetailRows.length + ' day' + (leaveDetailRows.length===1?'':'s') + ')',
    [
      { label:'Date', get:function(r){ return fmtDate_(r.date); } },
      { label:'Person', get:function(r){ return escHtml_(r.name); } },
      { label:'Type', get:function(r){ return escHtml_(r.type); } },
      { label:'Duration', get:function(r){ return r.fraction===0.5?'Half day':'Full day'; } },
      { label:'Purpose', wrap:true, get:function(r){ return escHtml_(r.purpose).replace(/\n/g,'<br>'); } },
      { label:'Status', get:function(r){ return escHtml_(r.status); } }
    ],
    leaveDetailRows, 'No leave records in this period.'
  );

  var dashUrl = dashboardUrl_();
  var ctaHTML = dashUrl ? (
    '<div style="text-align:center;margin-top:10px;">' +
      '<a href="' + dashUrl + '" target="_blank" style="display:inline-block;background:' + EMAIL_HEADER + ';color:#ffffff;text-decoration:none;font-size:13px;font-weight:700;letter-spacing:.3px;padding:13px 28px;border-radius:30px;box-shadow:0 6px 16px rgba(12,97,112,.28);">Open the full dashboard →</a>' +
    '</div>') : '';

  return '' +
  '<div style="background:' + EMAIL_BG + ';padding:24px 12px;">' +
  '<style>@keyframes digestFade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}} .digest-card{animation:digestFade .6s ease both}</style>' +
  '<div class="digest-card" style="font-family:\'Segoe UI\',Tahoma,Verdana,sans-serif;max-width:660px;margin:0 auto;background:#ffffff;border-radius:18px;overflow:hidden;box-shadow:0 12px 40px rgba(12,97,112,.14);color:#15323a;">' +
    '<div style="background:' + EMAIL_HEADER + ';padding:32px 30px 28px;text-align:center;color:#fff;">' +
      '<div style="font-weight:700;font-size:18px;letter-spacing:-.2px;">EMEA<span style="color:#bff3ef;">Social</span></div>' +
      '<h1 style="margin:14px 0 4px;font-size:24px;font-weight:500;font-family:Georgia,serif;">' + teamTitle + '</h1>' +
      '<div style="display:inline-block;margin-top:8px;background:rgba(255,255,255,.18);color:#ffffff;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;padding:5px 14px;border-radius:20px;">📊 ' + cadeLabel + ' Pulse</div>' +
      '<div style="margin-top:12px;font-size:12px;color:rgba(255,255,255,.72);">Reporting period · ' + win.label + '  ·  compared with ' + win.cmpLabel + '</div>' +
    '</div>' +
    '<div style="padding:28px 30px;">' +
      '<p style="font-size:16px;margin:0 0 4px;">Hi <strong>' + (regionLabel ? regionLabel + ' Team Lead' : 'Team') + '</strong>,</p>' +
      '<p style="color:#7d96a0;line-height:1.6;margin:0;">' + teamIntro + '</p>' +
      tiles +
      ratingBlock +
      agentTable +
      rejTable +
      extTable +
      topSummary +
      rejDetailTable +
      complianceTable +
      paDetailTable +
      checklistDetailTable +
      leaveDetailTable +
      '<h3 style="border-bottom:1px solid #e3edf2;padding-bottom:10px;font-weight:500;color:#3a5660;margin:28px 0 12px;font-family:Georgia,serif;">📊 Insights</h3>' +
      '<table width="100%" cellpadding="0" cellspacing="0">' + insightHTML + '</table>' +
      '<div style="margin-top:22px;background:rgba(20,145,155,.08);border-left:4px solid ' + accent + ';padding:16px 18px;border-radius:0 10px 10px 0;">' +
        '<div style="font-size:12px;font-weight:700;color:' + accent2 + ';text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">💡 Suggested actions</div>' +
        '<ul style="margin:0;padding-left:18px;font-size:13px;color:#3a5660;line-height:1.5;">' + sugHTML + '</ul>' +
      '</div>' +
      (regionLabel ? '' : discrepancyButtonHTML_(sentAtLabel)) +
      (regionLabel ? '' : ctaHTML) +
    '</div>' +
    '<div style="background:#f5f9fb;padding:18px;text-align:center;font-size:11px;color:#7d96a0;border-top:1px solid #e3edf2;">' +
      '<p style="margin:0;">Automated ' + cadeLabel + ' team digest · EMEA Social Posting Hub — Operations KPI Dashboard</p>' +
      '<p style="margin:5px 0 0;">A full downloadable PDF of this report is attached. © 2026 EMEA Social. All rights reserved.</p>' +
    '</div>' +
  '</div>' +
  '</div>';
}

/* ---------- PDF companion (downloadable, attached to every digest) ---------- */
function buildDigestPdf_(cadence, win, m, cmp, insSug, includeRatings, R, agentRows, data, regionLabel) {
  var name = 'Ops-KPI-' + (regionLabel ? regionLabel + '-' : '') + 'Team-Digest-' + cadence + '-' + win.label.replace(/[^\w-]+/g, '_');
  var doc = DocumentApp.create(name + ' (temp)');
  var body = doc.getBody();
  body.appendParagraph('EMEA Social — ' + (regionLabel ? regionLabel + ' Team' : 'Team') + ' Operations Digest').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph((cadence.charAt(0).toUpperCase()+cadence.slice(1)) + ' · ' + win.label + '  (vs ' + win.cmpLabel + ')');
  body.appendParagraph(' ');

  var mt = [['Metric','Current','Comparison']];
  function row(label, cur, prev){ mt.push([label, cur==null?'—':String(cur), prev==null?'—':String(prev)]); }
  row('Task Count', m.task, cmp.task);
  row('Internal Quality %', m.iq==null?null:m.iq.toFixed(1)+'%', cmp.iq==null?null:cmp.iq.toFixed(1)+'%');
  row('External Quality %', m.eq==null?null:m.eq.toFixed(1)+'%', cmp.eq==null?null:cmp.eq.toFixed(1)+'%');
  row('Internal Rejections', m.rej, cmp.rej);
  row('External Errors', m.extC, cmp.extC);
  row('Critical Misses', m.critC, cmp.critC);
  row('Utilization %', m.util.toFixed(1)+'%', cmp.util==null?null:cmp.util.toFixed(1)+'%');
  row('Production Hours', m.prodH.toFixed(1), cmp.prodH==null?null:cmp.prodH.toFixed(1));
  row('Production Utilization %', m.prodUtil==null?null:m.prodUtil.toFixed(1)+'%', cmp.prodUtil==null?null:cmp.prodUtil.toFixed(1)+'%');
  row('Non-Production Utilization %', m.nonProdUtil==null?null:m.nonProdUtil.toFixed(1)+'%', cmp.nonProdUtil==null?null:cmp.nonProdUtil.toFixed(1)+'%');
  row('Productive AHT', fmtAht_(m.ahtSec), fmtAht_(cmp.ahtSec));
  if (includeRatings) row('PKT (team avg) %', m.pkt==null?null:m.pkt.toFixed(1)+'%', cmp.pkt==null?null:cmp.pkt.toFixed(1)+'%');
  body.appendTable(mt);
  body.appendParagraph(' ');

  if (agentRows && agentRows.length) {
    body.appendParagraph('Individual Contributions').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var at = [['Agent','Tasks','Internal Quality','External','Critical','Utilization','Prod. Util','Non-Prod. Util','AHT']];
    agentRows.forEach(function(a){ at.push([a.name, String(a.task), a.iq==null?'—':a.iq.toFixed(1)+'%', String(a.extC), String(a.critC), a.util==null?'—':a.util.toFixed(1)+'%', a.prodUtil==null?'—':a.prodUtil.toFixed(1)+'%', a.nonProdUtil==null?'—':a.nonProdUtil.toFixed(1)+'%', fmtAht_(a.ahtSec)]); });
    body.appendTable(at);
    body.appendParagraph(' ');
  }

  body.appendParagraph('Internal Rejections — by category').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var rejCats = Object.keys(m.rejByCat);
  if (rejCats.length) {
    var rt = [['Category','Count']];
    rejCats.sort(function(a,b){ return m.rejByCat[b]-m.rejByCat[a]; }).forEach(function(c){ rt.push([c, String(m.rejByCat[c])]); });
    body.appendTable(rt);
  } else body.appendParagraph('No internal rejections logged this period.');
  body.appendParagraph(' ');

  // Full per-record detail sections — every row includes who it belongs to.
  function appendDetail_(title, rows, headers, mapper, emptyMsg) {
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (!rows.length) { body.appendParagraph(emptyMsg); body.appendParagraph(' '); return; }
    var t = [headers];
    rows.forEach(function(r){ t.push(mapper(r)); });
    body.appendTable(t);
    body.appendParagraph(' ');
  }

  appendDetail_('External Errors & Critical Misses — full detail',
    teamExtDetailRows_(data, win.start, win.end),
    ['Date','Person','Type','Task','Issue'],
    function(r){ return [fmtDate_(r.date), r.name, r.type, r.task, r.issue]; },
    'No external errors or critical misses logged this period.');

  appendDetail_('Internal Rejections — full detail',
    teamRejDetailRows_(data, win.start, win.end),
    ['Date','Person','Category','Count'],
    function(r){ return [fmtDate_(r.date), r.name, r.category, String(r.count)]; },
    'No internal rejections logged this period.');

  appendDetail_('Compliance — full detail',
    teamComplianceDetailRows_(data, win.start, win.end),
    ['Date','Person','Type','Details'],
    function(r){ return [fmtDate_(r.date), r.name, r.type, r.details]; },
    'No compliance entries logged this period.');

  appendDetail_('Process Adherence Misses (CILs) — full detail',
    teamPADetailRows_(data, win.start, win.end),
    ['Date','Person','Missed By','Columns Missed'],
    function(r){ return [fmtDate_(r.date), r.name, r.missedBy, r.columns]; },
    'No process-adherence misses logged this period.');

  appendDetail_('Checklist Misses — full detail',
    teamChecklistDetailRows_(data, win.start, win.end),
    ['Date','Person','Category'],
    function(r){ return [fmtDate_(r.date), r.name, r.category]; },
    'No checklist misses logged this period.');

  appendDetail_('Leaves — full detail',
    teamLeaveDetailRows_(data, win.start, win.end),
    ['Date','Person','Type','Duration','Purpose','Status'],
    function(r){ return [fmtDate_(r.date), r.name, r.type, r.fraction===0.5?'Half day':'Full day', r.purpose, r.status]; },
    'No leave records in this period.');

  if (includeRatings && R) {
    body.appendParagraph('Team Overall Rating: ' + R.overall.toFixed(2) + ' / 5.00 (' + overallLabel_(R.overall) + ')').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var rt2 = [['Parameter','Rating','Weight']];
    KPI_META_.forEach(function(o){ rt2.push([o.label, R.sc[o.k]+'/5 · '+RATING_LABELS[R.sc[o.k]], Math.round((KPI_WEIGHTS[o.k]||0)*100)+'%']); });
    body.appendTable(rt2);
    body.appendParagraph(' ');
  }

  body.appendParagraph('Insights').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  insSug.items.forEach(function(i){ body.appendListItem(i.text); });
  body.appendParagraph('Suggested Actions').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  insSug.suggestions.forEach(function(s){ body.appendListItem(s); });
  body.appendParagraph(' ');
  if (!regionLabel) body.appendParagraph('Spot a discrepancy? Raise it: ' + DISCREPANCY_FORM_URL + '  (SLA: 24 hours from this report)');

  doc.saveAndClose();
  var pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName(name + '.pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

/* ---------- send ---------- */
function sendTeamDigestEmail(cadence, testMode) {
  var data = getData();
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var win = digestWindow_(cadence, tz);
  var m = teamWindowMetrics_(data, win.start, win.end);
  var cmp = teamWindowMetrics_(data, win.cmpStart, win.cmpEnd);
  var includeRatings = (cadence === 'monthly' || cadence === 'quarterly');
  var R = includeRatings ? computeRatings_(data, null, m, { start: win.start, end: win.end }) : null;
  var agentRows = teamAgentRows_(data, win.start, win.end);
  var sentAtLabel = Utilities.formatDate(new Date(), tz, "d MMM yyyy, HH:mm") + ' IST';
  var insSug = buildTeamInsights_(m, cmp);
  var html = buildDigestHTML_(cadence, win, m, cmp, insSug, includeRatings, R, sentAtLabel, agentRows, data);
  var pdf = buildDigestPdf_(cadence, win, m, cmp, insSug, includeRatings, R, agentRows, data);
  var cadeLabel = cadence.charAt(0).toUpperCase() + cadence.slice(1);
  var subject = (testMode ? '[PREVIEW] ' : '') + '📊 ' + cadeLabel + ' Newsletter — ' + digestSubjectPeriod_(cadence, win, tz) + ' | EMEA SPH';
  var digestKey = 'digest-' + cadence;

  if (testMode) {
    MailApp.sendEmail({ to: DIGEST_PREVIEW_EMAIL, subject: subject, htmlBody: html, name: EMAIL_SENDER_NAME, attachments: [pdf] });
    return 'Preview sent to ' + DIGEST_PREVIEW_EMAIL + ' only (template review — the team was NOT emailed).';
  }

  if (digestAlreadySent_(digestKey, win.label))
    return 'Skipped — ' + cadence + ' team digest for ' + win.label + ' was already sent (see Email Status tab).';

  // Recipients come straight from the Team List tab (not from who happened to have activity).
  var teamEmails = teamListEmails_();
  if (!teamEmails.length) return 'Aborted: no emails found on the Team List tab.';
  MailApp.sendEmail({
    to: teamEmails.join(','), cc: DIGEST_PREVIEW_EMAIL,
    subject: subject, htmlBody: html, name: EMAIL_SENDER_NAME,
    attachments: [pdf]
  });
  logEmailStatus_(digestKey, false, win.label, [teamEmails.length + ' Team List member(s) (to) + cc ' + DIGEST_PREVIEW_EMAIL]);
  return 'Sent ' + cadence + ' team digest for ' + win.label + ' — 1 email to ' + teamEmails.length + ' Team List member(s), cc ' + DIGEST_PREVIEW_EMAIL + '.';
}

/* ============================================================
 *  NIGERIA-DEDICATED REPORT — daily · weekly · monthly · quarterly
 *  A team digest exactly like sendTeamDigestEmail() above, but scoped to
 *  only the Team List members tagged Country = "Nigeria" (Col C), sent to
 *  the Nigeria manager + SME instead of the whole team.
 *
 *  TEST FIRST — always previews to DIGEST_PREVIEW_EMAIL only:
 *     testNigeriaDigestDaily()   testNigeriaDigestWeekly()
 *     testNigeriaDigestMonthly()   testNigeriaDigestQuarterly()
 *  When happy, run installNigeriaDigestTriggers() once from the editor.
 * ============================================================ */

function testNigeriaDigestDaily()     { return sendNigeriaDigestEmail('daily', true); }
function testNigeriaDigestWeekly()    { return sendNigeriaDigestEmail('weekly', true); }
function testNigeriaDigestMonthly()   { return sendNigeriaDigestEmail('monthly', true); }
function testNigeriaDigestQuarterly() { return sendNigeriaDigestEmail('quarterly', true); }

function triggerNigeriaDigestDaily() { sendNigeriaDigestEmail('daily', false); }
function triggerNigeriaDigestWeekly() {
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  if (Utilities.formatDate(new Date(), tz, 'EEEE') !== 'Monday') return; // guard, in case the trigger fires off-schedule
  sendNigeriaDigestEmail('weekly', false);
}
function triggerNigeriaDigestMonthly() {
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  if (Utilities.formatDate(new Date(), tz, 'd') !== '2') return; // only actually sends on the 2nd
  sendNigeriaDigestEmail('monthly', false);
}
function triggerNigeriaDigestQuarterly() {
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var d = Utilities.formatDate(new Date(), tz, 'd'), mo = +Utilities.formatDate(new Date(), tz, 'M');
  if (d !== '2' || [1,4,7,10].indexOf(mo) < 0) return; // only fires on the 2nd of Jan/Apr/Jul/Oct
  sendNigeriaDigestEmail('quarterly', false);
}

/* Run ONCE from the editor to install the Nigeria daily/weekly/monthly/quarterly triggers.
 * Same time-zone caveat as installTeamDigestTriggers(): set the Apps Script
 * project time zone (⚙ Project Settings) to Asia/Kolkata first. */
function installNigeriaDigestTriggers() {
  ['triggerNigeriaDigestDaily', 'triggerNigeriaDigestWeekly', 'triggerNigeriaDigestMonthly', 'triggerNigeriaDigestQuarterly'].forEach(function(fn){
    ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
  });
  ScriptApp.newTrigger('triggerNigeriaDigestDaily').timeBased().everyDays(1).atHour(9).nearMinute(15).create();
  ScriptApp.newTrigger('triggerNigeriaDigestWeekly').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).nearMinute(15).create();
  ScriptApp.newTrigger('triggerNigeriaDigestMonthly').timeBased().everyDays(1).atHour(9).nearMinute(15).create();   // self-guards to the 2nd only
  ScriptApp.newTrigger('triggerNigeriaDigestQuarterly').timeBased().everyDays(1).atHour(9).nearMinute(15).create(); // self-guards to the 2nd of Jan/Apr/Jul/Oct
  return 'Installed. Nigeria digest — Daily/Weekly(Mon)/Monthly(2nd)/Quarterly(2nd of Jan·Apr·Jul·Oct) @~9:15am — confirm the project time zone is Asia/Kolkata first.';
}

function sendNigeriaDigestEmail(cadence, testMode) {
  var fullData = getData();
  var tz = ss_().getSpreadsheetTimeZone() || 'Asia/Kolkata';
  var idxSet = regionIdxSet_(fullData, NIGERIA_COUNTRY_VALUE);
  if (!Object.keys(idxSet).length)
    return 'Aborted: no Team List member is tagged Country = "' + NIGERIA_COUNTRY_VALUE + '" (Col C).';
  var data = filterDataToRegion_(fullData, idxSet);

  var win = digestWindow_(cadence, tz);
  var m = teamWindowMetrics_(data, win.start, win.end);
  var cmp = teamWindowMetrics_(data, win.cmpStart, win.cmpEnd);
  var includeRatings = (cadence === 'monthly' || cadence === 'quarterly');
  var R = includeRatings ? computeRatings_(data, null, m, { start: win.start, end: win.end }) : null;
  var agentRows = teamAgentRows_(data, win.start, win.end);
  var sentAtLabel = Utilities.formatDate(new Date(), tz, "d MMM yyyy, HH:mm") + ' IST';
  var insSug = buildTeamInsights_(m, cmp);
  var html = buildDigestHTML_(cadence, win, m, cmp, insSug, includeRatings, R, sentAtLabel, agentRows, data, NIGERIA_COUNTRY_VALUE);
  var pdf = buildDigestPdf_(cadence, win, m, cmp, insSug, includeRatings, R, agentRows, data, NIGERIA_COUNTRY_VALUE);
  var cadeLabel = cadence.charAt(0).toUpperCase() + cadence.slice(1);
  var subject = (testMode ? '[PREVIEW] ' : '') + '📊 ' + NIGERIA_COUNTRY_VALUE + ' ' + cadeLabel + ' Performance — ' + digestSubjectPeriod_(cadence, win, tz) + ' | EMEA SPH';
  var digestKey = 'digest-nigeria-' + cadence;

  if (testMode) {
    MailApp.sendEmail({ to: DIGEST_PREVIEW_EMAIL, subject: subject, htmlBody: html, name: EMAIL_SENDER_NAME, attachments: [pdf] });
    return 'Preview sent to ' + DIGEST_PREVIEW_EMAIL + ' only (template review — the manager/SME were NOT emailed).';
  }

  if (digestAlreadySent_(digestKey, win.label))
    return 'Skipped — ' + cadence + ' Nigeria report for ' + win.label + ' was already sent (see Email Status tab).';

  var toList = NIGERIA_REPORT_TO.filter(Boolean), ccList = NIGERIA_REPORT_CC.filter(Boolean);
  if (!toList.length) return 'Aborted: NIGERIA_REPORT_TO has no addresses configured.';
  MailApp.sendEmail({
    to: toList.join(','), cc: ccList.join(','),
    subject: subject, htmlBody: html, name: EMAIL_SENDER_NAME,
    attachments: [pdf]
  });
  logEmailStatus_(digestKey, false, win.label, [toList.length + ' to (' + toList.join('; ') + ') + ' + ccList.length + ' cc (' + ccList.join('; ') + ')']);
  return 'Sent ' + cadence + ' Nigeria report for ' + win.label + ' — to ' + toList.join(', ') + ', cc ' + ccList.join(', ') + '.';
}

/* ---------- diagnostics (run from the editor) ---------- */
function findSplits() {
  var qs = readSheet_(TABS.queue), rs = readSheet_(TABS.rej), ts = readSheet_(TABS.team),
      us = readSheet_(TABS.util), es = readSheet_(TABS.ext);
  var ID = buildIdentity_(qs, rs, ts);
  var out = ['Auto-linked email spellings (wrong -> canonical):'];
  var ks = Object.keys(ID.aliasMap);
  if (!ks.length) out.push('   (none — every email already matched the Team List)');
  ks.forEach(function(k){ out.push('   ' + k + '   ->   ' + ID.aliasMap[k]); });
  // emails still unknown to Team List and NOT auto-linked
  out.push('\nEmails seen in data but NOT in Team List and NOT auto-linked:');
  var seen = {};
  [[us,'Utilization'],[es,'External Errors'],[qs,'Queue Tracker'],[rs,'Rejection Categories']].forEach(function(p){
    p[0].rows.forEach(function(r){
      var raw = normEmail_(col_(p[0], r, 'Email')); if (!raw) return;
      var c = ID.canon(raw);
      if (!ID.teamEmails[c]) { var key = raw + ' (' + p[1] + ')'; if (!seen[key]) { seen[key] = true; out.push('   ' + key); } }
    });
  });
  if (Object.keys(seen).length === 0) out.push('   (none)');
  Logger.log(out.join('\n')); return out.join('\n');
}
function diagnose(substr) {
  substr = String(substr || '').toLowerCase();
  var ss = ss_(); var out = ['Searching for: "' + substr + '"   (taskvol = Assigned By + Trafficking + Live QC)'];
  [TABS.queue, TABS.rej, TABS.team, TABS.util, TABS.ext].forEach(function(tab){
    var sh = ss.getSheetByName(tab); out.push('\n== ' + tab + ' ==');
    if (!sh) { out.push('   (not found)'); return; }
    var vals = sh.getDataRange().getValues(); if (vals.length < 2) { out.push('   (empty)'); return; }
    var hl = vals[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var ei = hl.indexOf('email'); if (ei < 0) { out.push('   (no Email column)'); return; }
    var ni = hl.indexOf('name'); if (ni < 0) ni = hl.indexOf('name (filtered)');
    var isQ = (tab === TABS.queue);
    var ai = hl.indexOf('assigned by'), ti = hl.indexOf('trafficking'), qi = hl.indexOf('live qc');
    var agg = {};
    for (var i = 1; i < vals.length; i++) {
      var raw = vals[i][ei]; if (raw === '' || raw == null) continue;
      var s = String(raw), nm = ni >= 0 ? String(vals[i][ni] || '') : '';
      if (substr && (s + ' ' + nm).toLowerCase().indexOf(substr) < 0) continue;
      var key = JSON.stringify(s) + (nm ? '  name=' + JSON.stringify(nm) : '') + '  norm=' + normEmail_(s);
      if (!agg[key]) agg[key] = { n: 0, vol: 0 };
      agg[key].n++;
      if (isQ) agg[key].vol += num_(vals[i][ai]) + num_(vals[i][ti]) + num_(vals[i][qi]);
    }
    var keys = Object.keys(agg);
    if (!keys.length) out.push('   (no match)');
    else keys.forEach(function(k){ out.push('   x' + agg[k].n + (isQ ? '  taskvol=' + agg[k].vol : '') + '   ' + k); });
  });
  Logger.log(out.join('\n')); return out.join('\n');
}

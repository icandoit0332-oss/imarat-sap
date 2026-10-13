// ══════════════════════════════════════════════════════════════════
//  IMARAT SAP RMS — CHANGE REQUEST FORM (CRF) MODULE
//  Add this as a NEW file (CRF.gs) in the same Apps Script project.
//  Then add ONE line at the top of the existing doGet AND doPost:
//
//      var crf = crfRoute(e); if (crf) return crf;
//
//  Flow:  SAP team raises CRF in dashboard  ->  reviewer (Abdul) is emailed
//         reviewer returns it, or sends it to the approver (Saad Sb)
//         approver approves / rejects from the email  ->  raiser + reviewer emailed
//         approved CRF is marked Implemented with transport / reference
//  Data:  sheet "CRF" in the same spreadsheet (created automatically)
// ══════════════════════════════════════════════════════════════════

var CRF_CFG = {
  sheetName:      'CRF',
  approverName:   'Saad Khalid',
  approverEmail:  'saad.khalid@imarat.com.pk',
  reviewerName:   'Abdul Wasay',
  reviewerEmail:  'abdul.wasay@imarat.com.pk',
  execUrl:        'https://script.google.com/macros/s/AKfycbxGsPIkGoZxVFqZ-17h1TLm1wAEB_h0J8Dax2DBlTNQgN6nrCVqJUjt7fZaIgMZWowg/exec',
  dashboardUrl:   'https://icandoit0332-oss.github.io/imarat-sap/team.html',
  firebaseApiKey: 'AIzaSyBei-lo6GIehz3UT93xg2KTs5bU_1-4osQ',
  firebaseProject:'imarat-sap-aff7f'
};

var CRF_HEADERS = ['CRF ID','Raised On','Raised By','Raiser Email','Module','Change Type','Title',
  'Description','Business Reason','Impact','Company Codes','Linked REQ','Urgency','Risk',
  'Test Plan','Rollback Plan','Planned Date','Status','Reviewer Note','Sent On','Approval Token',
  'Decision','Decision On','Approver Remarks','Transport Ref','Implemented On','History','Last Updated'];

// form field -> column header
var CRF_FIELDS = { module:'Module', changeType:'Change Type', title:'Title', description:'Description',
  reason:'Business Reason', impact:'Impact', companyCodes:'Company Codes', linkedReq:'Linked REQ',
  urgency:'Urgency', risk:'Risk', testPlan:'Test Plan', rollback:'Rollback Plan', plannedDate:'Planned Date' };
var CRF_TEXT_COLS = ['Company Codes','Linked REQ','Planned Date','Transport Ref'];
var CRF_REQUIRED = ['module','changeType','title','description','reason','impact','urgency','risk','rollback'];

// ── ROUTER ──────────────────────────────────────────────────────────
function crfRoute(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = String(p.action || '');
  if (action.indexOf('crf') !== 0) return null;

  try {
    // Public, token-protected pages opened from the approver's email
    if (action === 'crfDecide')       return crfDecidePage_(p);
    if (action === 'crfDecideSubmit') return crfDecideSubmit_(p);

    // Everything else needs a signed-in dashboard user
    var me = crfAuth_(p.idToken);
    if (!me) return crfJson_({ success:false, error:'Not authorised — sign in again.' });
    if (me.role !== 'admin' && me.role !== 'editor') return crfJson_({ success:false, error:'Your role cannot use CRFs.' });

    if (action === 'crfRead')      return crfJson_(crfRead_(me));
    if (action === 'crfSubmit')    return crfJson_(crfSubmit_(p, me));
    if (action === 'crfReturn')    return crfJson_(crfReturn_(p, me));
    if (action === 'crfSend')      return crfJson_(crfSend_(p, me, false));
    if (action === 'crfRemind')    return crfJson_(crfSend_(p, me, true));
    if (action === 'crfImplement') return crfJson_(crfImplement_(p, me));
    return crfJson_({ success:false, error:'Unknown CRF action.' });
  } catch (err) {
    Logger.log('CRF error: ' + err.stack);
    if (action === 'crfDecide' || action === 'crfDecideSubmit') return crfPage_('Error', '<p>' + crfEsc_(err.message) + '</p>', '#DC2626');
    return crfJson_({ success:false, error: err.message });
  }
}

// ── AUTH: verify Firebase ID token, read role from Firestore users/{uid} ──
function crfAuth_(idToken) {
  if (!idToken) return null;
  var cache = CacheService.getScriptCache();
  var key = 'crfauth_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + CRF_CFG.firebaseApiKey, {
    method:'post', contentType:'application/json', payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions:true });
  if (r.getResponseCode() !== 200) return null;
  var u = (JSON.parse(r.getContentText()).users || [])[0];
  if (!u || !u.email) return null;

  var d = UrlFetchApp.fetch('https://firestore.googleapis.com/v1/projects/' + CRF_CFG.firebaseProject +
    '/databases/(default)/documents/users/' + u.localId, {
    headers: { Authorization: 'Bearer ' + idToken }, muteHttpExceptions:true });
  var role = 'viewer', name = u.email;
  if (d.getResponseCode() === 200) {
    var f = JSON.parse(d.getContentText()).fields || {};
    role = (f.role && f.role.stringValue) || 'viewer';
    name = (f.name && f.name.stringValue) || u.email;
  }
  var me = { uid:u.localId, email:String(u.email).toLowerCase(), name:name, role:role };
  cache.put(key, JSON.stringify(me), 600);
  return me;
}

// ── SHEET HELPERS ───────────────────────────────────────────────────
function crfSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CRF_CFG.sheetName);
  if (!sh) {
    sh = ss.insertSheet(CRF_CFG.sheetName);
    sh.appendRow(CRF_HEADERS);
    sh.getRange(1, 1, 1, CRF_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function crfCol_(h) { return CRF_HEADERS.indexOf(h) + 1; }

function crfAll_() {
  var sh = crfSheet_();
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var o = { _row: i + 1 };
    for (var c = 0; c < CRF_HEADERS.length; c++) o[CRF_HEADERS[c]] = v[i][c] instanceof Date ? v[i][c].toISOString() : v[i][c];
    out.push(o);
  }
  return out;
}
function crfFind_(id) {
  var all = crfAll_();
  for (var i = 0; i < all.length; i++) if (all[i]['CRF ID'] === id) return all[i];
  return null;
}
function crfSet_(rec, changes) {
  var sh = crfSheet_();
  changes['Last Updated'] = new Date().toISOString();
  for (var h in changes) {
    var cell = sh.getRange(rec._row, crfCol_(h));
    if (CRF_TEXT_COLS.indexOf(h) >= 0) cell.setNumberFormat('@');
    cell.setValue(changes[h]); rec[h] = changes[h];
  }
}
function crfHistory_(rec, by, act, note) {
  var h = [];
  try { h = JSON.parse(rec['History'] || '[]'); } catch (x) {}
  h.push({ t: new Date().toISOString(), by: by, a: act, n: note || '' });
  return JSON.stringify(h);
}
function crfNextId_() {
  var yr = new Date().getFullYear(), max = 0, pre = 'CRF-' + yr + '-';
  crfAll_().forEach(function (r) {
    var id = String(r['CRF ID']);
    if (id.indexOf(pre) === 0) max = Math.max(max, parseInt(id.slice(pre.length), 10) || 0);
  });
  return pre + ('00' + (max + 1)).slice(-3);
}
function crfLocked_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function crfPublic_(r) {
  var o = {};
  for (var h in r) if (h !== 'Approval Token' && h !== '_row') o[h] = r[h];
  return o;
}

// ── ACTIONS ─────────────────────────────────────────────────────────
function crfRead_(me) {
  return { success:true, me:{ email:me.email, role:me.role, name:me.name }, data: crfAll_().map(crfPublic_) };
}

function crfSubmit_(p, me) {
  CRF_REQUIRED.forEach(function (f) { if (!String(p[f] || '').trim()) throw new Error('Missing: ' + CRF_FIELDS[f]); });
  return crfLocked_(function () {
    var now = new Date().toISOString(), rec;
    if (p.crfId) {
      rec = crfFind_(p.crfId);
      if (!rec) throw new Error('CRF not found.');
      if (rec['Status'] !== 'Submitted' && rec['Status'] !== 'Returned') throw new Error('Only Submitted or Returned CRFs can be edited.');
      if (me.role !== 'admin' && String(rec['Raiser Email']).toLowerCase() !== me.email) throw new Error('Only the raiser can edit this CRF.');
      var ch = { 'Status':'Submitted', 'History': crfHistory_(rec, me.name, rec['Status'] === 'Returned' ? 'Resubmitted' : 'Edited', '') };
      for (var f in CRF_FIELDS) ch[CRF_FIELDS[f]] = String(p[f] || '').trim();
      crfSet_(rec, ch);
    } else {
      var id = crfNextId_();
      var row = CRF_HEADERS.map(function () { return ''; });
      row[0] = id; row[1] = now; row[2] = me.name; row[3] = me.email;
      for (var g in CRF_FIELDS) row[crfCol_(CRF_FIELDS[g]) - 1] = String(p[g] || '').trim();
      row[crfCol_('Status') - 1] = 'Submitted';
      row[crfCol_('History') - 1] = JSON.stringify([{ t:now, by:me.name, a:'Raised', n:'' }]);
      row[crfCol_('Last Updated') - 1] = now;
      var sh = crfSheet_();
      sh.appendRow(row);
      // keep text columns as text (stop Sheets turning "1100,1200" or dates into numbers/dates)
      var last = sh.getLastRow();
      CRF_TEXT_COLS.forEach(function (h) { sh.getRange(last, crfCol_(h)).setNumberFormat('@').setValue(row[crfCol_(h) - 1]); });
      rec = crfFind_(id);
    }
    if (me.email !== CRF_CFG.reviewerEmail.toLowerCase()) {
      crfMail_(CRF_CFG.reviewerEmail, '', rec['CRF ID'] + ' submitted for review — ' + rec['Title'],
        crfTpl_('Change Request — Review Required',
          '<p style="margin:0 0 14px">' + crfEsc_(rec['Raised By']) + ' has ' + (p.crfId ? 'updated' : 'raised') + ' a change request. Review it in the dashboard and send it for approval or return it.</p>'
          + crfTable_(rec) + crfButton_(CRF_CFG.dashboardUrl, 'Open CRF register', '#14432A')));
    }
    return { success:true, id: rec['CRF ID'] };
  });
}

function crfReturn_(p, me) {
  if (me.role !== 'admin') throw new Error('Only the reviewer can return a CRF.');
  var note = String(p.note || '').trim();
  if (!note) throw new Error('Give a reason for returning the CRF.');
  return crfLocked_(function () {
    var rec = crfFind_(p.id);
    if (!rec) throw new Error('CRF not found.');
    if (rec['Status'] !== 'Submitted') throw new Error('Only Submitted CRFs can be returned.');
    crfSet_(rec, { 'Status':'Returned', 'Reviewer Note':note, 'History': crfHistory_(rec, me.name, 'Returned', note) });
    crfMail_(rec['Raiser Email'], '', rec['CRF ID'] + ' returned for changes — ' + rec['Title'],
      crfTpl_('Change Request Returned',
        '<p style="margin:0 0 14px">Your change request has been returned. Update it in the dashboard and resubmit.</p>'
        + crfNote_('Reviewer note', note, '#9A7B2F') + crfTable_(rec) + crfButton_(CRF_CFG.dashboardUrl, 'Open CRF register', '#14432A')));
    return { success:true };
  });
}

function crfSend_(p, me, isReminder) {
  if (me.role !== 'admin') throw new Error('Only the reviewer can send a CRF for approval.');
  return crfLocked_(function () {
    var rec = crfFind_(p.id);
    if (!rec) throw new Error('CRF not found.');
    var ch = {};
    if (isReminder) {
      if (rec['Status'] !== 'Pending Approval') throw new Error('Reminders apply only to CRFs pending approval.');
      ch['History'] = crfHistory_(rec, me.name, 'Reminder sent', '');
    } else {
      if (rec['Status'] !== 'Submitted') throw new Error('Only Submitted CRFs can be sent for approval.');
      var note = String(p.note || '').trim();
      ch = { 'Status':'Pending Approval', 'Sent On': new Date().toISOString(), 'Approval Token': Utilities.getUuid(),
             'Reviewer Note': note, 'History': crfHistory_(rec, me.name, 'Sent for approval', note) };
    }
    crfSet_(rec, ch);

    var base = CRF_CFG.execUrl + '?action=crfDecide&id=' + encodeURIComponent(rec['CRF ID']) + '&token=' + rec['Approval Token'];
    var intro = (isReminder ? '<p style="margin:0 0 10px;color:#9A7B2F;font-weight:bold">Reminder — this request is awaiting your decision.</p>' : '')
      + '<p style="margin:0 0 14px">The change request below has been reviewed and is submitted for your approval before implementation.</p>';
    var body = intro + (rec['Reviewer Note'] ? crfNote_('Reviewer note', rec['Reviewer Note'], '#14432A') : '') + crfTable_(rec)
      + '<table role="presentation" style="width:100%;margin-top:18px"><tr>'
      + '<td style="padding-right:6px">' + crfButton_(base + '&d=approve', 'Approve', '#14432A', true) + '</td>'
      + '<td style="padding-left:6px">'  + crfButton_(base + '&d=reject',  'Reject',  '#B42318', true) + '</td>'
      + '</tr></table>';
    var subj = (isReminder ? 'Reminder: ' : '') + rec['CRF ID'] + ' for approval — ' + rec['Module'] + ' — ' + rec['Title'];
    crfMail_(CRF_CFG.approverEmail, '', subj, crfTpl_('Change Request — Approval Required', body));

    // Reviewer's copy carries no decision links, so only the approver can decide
    crfMail_(CRF_CFG.reviewerEmail, '', 'Copy: ' + subj, crfTpl_('Change Request — Sent to ' + CRF_CFG.approverName,
      '<p style="margin:0 0 14px">' + (isReminder ? 'Reminder sent' : 'Sent') + ' to ' + crfEsc_(CRF_CFG.approverName) + ' (' + crfEsc_(CRF_CFG.approverEmail) + ') for approval.</p>'
      + (rec['Reviewer Note'] ? crfNote_('Reviewer note', rec['Reviewer Note'], '#14432A') : '') + crfTable_(rec)));
    return { success:true };
  });
}

function crfImplement_(p, me) {
  var ref = String(p.transportRef || '').trim();
  if (!ref) throw new Error('Enter the transport request or reference.');
  return crfLocked_(function () {
    var rec = crfFind_(p.id);
    if (!rec) throw new Error('CRF not found.');
    if (rec['Status'] !== 'Approved') throw new Error('Only Approved CRFs can be marked implemented.');
    if (me.role !== 'admin' && String(rec['Raiser Email']).toLowerCase() !== me.email) throw new Error('Only the raiser or reviewer can mark this implemented.');
    crfSet_(rec, { 'Status':'Implemented', 'Transport Ref':ref, 'Implemented On': new Date().toISOString(),
                   'History': crfHistory_(rec, me.name, 'Implemented', ref) });
    crfMail_(CRF_CFG.reviewerEmail, '', rec['CRF ID'] + ' implemented — ' + rec['Title'],
      crfTpl_('Change Request Implemented', '<p style="margin:0 0 14px">Implemented by ' + crfEsc_(me.name) + '. Reference: <b>' + crfEsc_(ref) + '</b></p>' + crfTable_(rec)));
    return { success:true };
  });
}

// ── APPROVER PAGES (opened from email) ──────────────────────────────
// The email links open a confirmation page; nothing is recorded until the
// approver presses the button (protects against mail scanners pre-opening links).
function crfDecidePage_(p) {
  var rec = crfFind_(p.id);
  if (!rec || !p.token || rec['Approval Token'] !== p.token) {
    if (rec && rec['Decision']) return crfPage_('Decision already recorded', '<p>' + crfEsc_(rec['CRF ID']) + ' was <b>' + crfEsc_(rec['Decision']) + '</b> on ' + crfFmt_(rec['Decision On']) + '.</p>', '#14432A');
    return crfPage_('Link not valid', '<p>This approval link is no longer valid.</p>', '#B42318');
  }
  var approve = p.d !== 'reject';
  var color = approve ? '#14432A' : '#B42318';
  // The form is submitted with google.script.run and the page updates in place.
  // (A normal form submit would navigate inside Google's sandbox frame, which
  //  script.google.com refuses to load — the "refused to connect" screen.)
  var form = '<form id="crfForm" onsubmit="return crfGo(event)">'
    + '<label style="display:block;font-weight:bold;margin:16px 0 6px">Remarks' + (approve ? ' (optional)' : ' — reason for rejection *') + '</label>'
    + '<textarea id="crfRemarks" ' + (approve ? '' : 'required') + ' style="width:100%;box-sizing:border-box;min-height:90px;border:1px solid #C9D6CC;border-radius:6px;padding:10px;font-family:inherit;font-size:14px"></textarea>'
    + '<button id="crfBtn" type="submit" style="width:100%;margin-top:12px;background:' + color + ';color:#fff;border:0;border-radius:6px;padding:13px;font-size:15px;font-weight:bold;cursor:pointer">'
    + (approve ? 'Confirm Approval' : 'Confirm Rejection') + '</button>'
    + '<div id="crfErr" style="display:none;margin-top:10px;color:#B42318;font-size:13px"></div></form>'
    + '<script>'
    + 'var CRF_P=' + JSON.stringify({ id: rec['CRF ID'], token: p.token, d: approve ? 'approve' : 'reject' }).replace(/</g, '\\u003c') + ';'
    + 'function crfGo(e){e.preventDefault();var b=document.getElementById("crfBtn"),r=document.getElementById("crfRemarks").value.trim();'
    + 'if(CRF_P.d==="reject"&&!r){crfErr("Please give a reason for rejection.");return false;}'
    + 'b.disabled=true;b.style.opacity="0.6";b.textContent="Recording decision…";'
    + 'google.script.run.withSuccessHandler(function(res){document.getElementById("crfTitle").textContent=res.title;'
    + 'document.getElementById("crfTitle").style.color=res.color;document.getElementById("crfContent").innerHTML=res.html;})'
    + '.withFailureHandler(function(err){b.disabled=false;b.style.opacity="1";b.textContent="Try again";crfErr("Could not record the decision: "+(err&&err.message||err));})'
    + '.crfDecideRun({id:CRF_P.id,token:CRF_P.token,d:CRF_P.d,remarks:r});return false;}'
    + 'function crfErr(m){var x=document.getElementById("crfErr");x.textContent=m;x.style.display="block";}'
    + '</script>';
  return crfPage_((approve ? 'Approve ' : 'Reject ') + rec['CRF ID'], crfTable_(rec) + form, color);
}

// Called from the confirm page via google.script.run (must stay public — no trailing underscore)
function crfDecideRun(p) {
  var r = crfDecide_(p || {});
  return { title: r.title, html: r.html, color: r.color };
}

// Older emails / direct GET submissions still work
function crfDecideSubmit_(p) {
  var r = crfDecide_(p);
  return crfPage_(r.title, r.html, r.color);
}

function crfDecide_(p) {
  var approve = p.d !== 'reject';
  var remarks = String(p.remarks || '').trim();
  if (!approve && !remarks) return { title:'Reason required', html:'<p>Please give a reason for rejection.</p>', color:'#B42318' };
  var rec = crfLocked_(function () {
    var r = crfFind_(p.id);
    if (!r || !p.token || r['Approval Token'] !== p.token || r['Status'] !== 'Pending Approval') return null;
    var decision = approve ? 'Approved' : 'Rejected';
    crfSet_(r, { 'Status':decision, 'Decision':decision, 'Decision On': new Date().toISOString(),
                 'Approver Remarks':remarks, 'Approval Token':'', 'History': crfHistory_(r, CRF_CFG.approverName, decision, remarks) });
    return r;
  });
  if (!rec) {
    var cur = crfFind_(p.id);
    if (cur && cur['Decision']) return { title:'Decision already recorded', color:'#14432A',
      html:'<p>' + crfEsc_(cur['CRF ID']) + ' was <b>' + crfEsc_(cur['Decision']) + '</b> on ' + crfFmt_(cur['Decision On']) + '.</p>' };
    return { title:'Link not valid', html:'<p>This approval link is no longer valid.</p>', color:'#B42318' };
  }

  var dc = approve ? '#14432A' : '#B42318';
  var body = '<p style="margin:0 0 14px">' + crfEsc_(CRF_CFG.approverName) + ' has <b style="color:' + dc + '">' + rec['Decision'].toLowerCase() + '</b> this change request'
    + (approve ? '. It may now be implemented; mark it Implemented in the dashboard with the transport reference.' : '. It must not be implemented.') + '</p>'
    + (remarks ? crfNote_('Approver remarks', remarks, dc) : '') + crfTable_(rec);
  var subj = rec['CRF ID'] + ' ' + rec['Decision'].toLowerCase() + ' — ' + rec['Title'];
  var raiser = String(rec['Raiser Email'] || '').toLowerCase();
  var rev = CRF_CFG.reviewerEmail.toLowerCase();
  crfMail_(raiser || rev, (raiser && raiser !== rev) ? rev : '', subj, crfTpl_('Change Request ' + rec['Decision'], body));

  var done = '<div style="text-align:center;padding:8px 0 4px">'
    + '<div style="width:64px;height:64px;border-radius:50%;margin:0 auto 14px;display:flex;align-items:center;justify-content:center;font-size:32px;color:#fff;background:' + dc + '">' + (approve ? '&#10003;' : '&#10005;') + '</div>'
    + '<p style="font-size:15px;margin:0 0 6px"><b>' + crfEsc_(rec['CRF ID']) + '</b> &mdash; ' + crfEsc_(rec['Title']) + '</p>'
    + '<p style="margin:0 0 14px">' + (approve ? 'Approved' : 'Rejected') + ' on ' + crfFmt_(rec['Decision On']) + '.</p>'
    + (remarks ? '<div style="text-align:left">' + crfNote_('Your remarks', remarks, dc) + '</div>' : '')
    + '<p style="color:#6B7F72;font-size:13px;margin:14px 0 0">' + crfEsc_(CRF_CFG.reviewerName) + ' and the raiser have been notified by email. You can close this page.</p></div>';
  return { title: approve ? 'Change request approved' : 'Change request rejected', html: done, color: dc };
}

function authorizeCRF() {
  UrlFetchApp.fetch('https://www.google.com');
  Logger.log('CRF permissions OK');
}

// ── OUTPUT / EMAIL HELPERS ──────────────────────────────────────────
function crfJson_(o) { o.crf = true; return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function crfMail_(to, cc, subj, html) {
  try {
    if (!to) return;
    var o = { to: to, subject: subj, htmlBody: html, name: 'IMARAT SAP — Change Control' };
    if (cc) o.cc = cc;
    MailApp.sendEmail(o);
  } catch (e) { Logger.log('CRF mail error: ' + e.message); }
}

function crfEsc_(s) {
  return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function crfFmt_(iso) {
  if (!iso) return '';
  var d = new Date(iso); if (isNaN(d)) return crfEsc_(iso);
  var n = d.getDate(), s = (n % 10 === 1 && n !== 11) ? 'st' : (n % 10 === 2 && n !== 12) ? 'nd' : (n % 10 === 3 && n !== 13) ? 'rd' : 'th';
  var m = ['January','February','March','April','May','June','July','August','September','October','November','December'][d.getMonth()];
  return n + s + ' ' + m + ', ' + d.getFullYear();
}

function crfTpl_(title, content) {
  return '<div style="font-family:Calibri,Arial,sans-serif;max-width:640px;margin:0 auto;background:#ffffff;color:#1F2A24">'
    + '<div style="padding:18px 24px 10px;border-bottom:3px solid #9A7B2F">'
    + '<div style="font-variant:small-caps;letter-spacing:1px;font-size:18px;font-weight:bold;color:#14432A">Imarat Group of Companies</div>'
    + '<div style="font-size:12px;color:#9A7B2F;margin-top:2px">SAP Programme · Change Control</div></div>'
    + '<div style="padding:22px 24px">'
    + '<h2 style="color:#14432A;font-size:18px;margin:0 0 14px">' + crfEsc_(title) + '</h2>'
    + '<div style="font-size:14px;line-height:1.5">' + content + '</div>'
    + '<div style="margin-top:22px;padding-top:12px;border-top:1px solid #E9F0E8;font-size:11px;color:#6B7F72">Automated notification from the SAP Requirement Management System.</div>'
    + '</div></div>';
}

function crfTable_(r) {
  var rows = [
    ['CRF ID', '<b style="color:#14432A">' + crfEsc_(r['CRF ID']) + '</b>'],
    ['Title', '<b>' + crfEsc_(r['Title']) + '</b>'],
    ['Module', crfEsc_(r['Module'])],
    ['Change type', crfEsc_(r['Change Type'])],
    ['Urgency', crfEsc_(r['Urgency'])],
    ['Risk', crfEsc_(r['Risk'])],
    ['Raised by', crfEsc_(r['Raised By']) + ' · ' + crfFmt_(r['Raised On'])],
    ['Company codes', crfEsc_(r['Company Codes'])],
    ['Linked requirement', crfEsc_(r['Linked REQ'])],
    ['Planned date', crfFmt_(r['Planned Date'])],
    ['Description', crfEsc_(r['Description']).replace(/\n/g, '<br>')],
    ['Business reason', crfEsc_(r['Business Reason']).replace(/\n/g, '<br>')],
    ['Impact / systems', crfEsc_(r['Impact']).replace(/\n/g, '<br>')],
    ['Test plan', crfEsc_(r['Test Plan']).replace(/\n/g, '<br>')],
    ['Rollback plan', crfEsc_(r['Rollback Plan']).replace(/\n/g, '<br>')]
  ].filter(function (x) { return x[1] !== ''; });
  return '<table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid #DDE6DF">'
    + rows.map(function (x, i) {
      var bg = i % 2 ? '#F5F8F4' : '#ffffff';
      return '<tr><td style="padding:8px 10px;width:140px;vertical-align:top;background:#E9F0E8;color:#14432A;font-weight:bold">' + x[0] + '</td>'
        + '<td style="padding:8px 10px;vertical-align:top;background:' + bg + '">' + x[1] + '</td></tr>';
    }).join('') + '</table>';
}

function crfNote_(label, text, color) {
  return '<div style="margin:0 0 14px;padding:10px 12px;border-left:3px solid ' + color + ';background:#F5F8F4">'
    + '<div style="font-size:11px;font-weight:bold;color:' + color + ';text-transform:uppercase;margin-bottom:4px">' + label + '</div>'
    + '<div>' + crfEsc_(text).replace(/\n/g, '<br>') + '</div></div>';
}

function crfButton_(url, label, color, block) {
  return '<a href="' + url + '" style="' + (block ? 'display:block;' : 'display:inline-block;margin-top:16px;') + 'background:' + color
    + ';color:#ffffff;text-align:center;padding:12px 22px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:14px">' + label + '</a>';
}

function crfPage_(title, content, color) {
  return HtmlService.createHtmlOutput('<html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;background:#F5F8F4;font-family:Calibri,Arial,sans-serif;color:#1F2A24">'
    + '<div style="max-width:620px;margin:24px auto;background:#fff;border:1px solid #DDE6DF;border-radius:8px;overflow:hidden">'
    + '<div style="padding:16px 22px 8px;border-bottom:3px solid #9A7B2F"><div style="font-variant:small-caps;letter-spacing:1px;font-size:17px;font-weight:bold;color:#14432A">Imarat Group of Companies</div>'
    + '<div style="font-size:12px;color:#9A7B2F">SAP Programme · Change Control</div></div>'
    + '<div style="padding:20px 22px"><h2 id="crfTitle" style="margin:0 0 14px;color:' + color + ';font-size:19px">' + crfEsc_(title) + '</h2><div id="crfContent">' + content + '</div></div>'
    + '</div></body></html>').setTitle('IMARAT — Change Request');
}

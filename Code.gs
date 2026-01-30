// --- CONFIGURATION ---
var SPREADSHEET_ID = '1RO1tTij8KY872eigB0Wm4tRqBExYKHj2p2SHd8WKy5c';
var ALLOWED_DOMAIN = '@portlandmaine.gov';

// --- SHEET NAMES ---
var FORM_SHEET_NAME = 'FROI Form';
var ADMIN_SHEET_NAME = 'Admins';
var DOCS_SHEET_NAME = 'Case_Docs';
var CONTACTS_SHEET_NAME = 'Case_Contacts';
var RESTR_SHEET_NAME = 'Case_Restrictions';
var LOSTTIME_SHEET_NAME = 'Case_LostTime';
var COMMLOG_SHEET_NAME = 'CommLog';
var DEPTS_SHEET_NAME = 'Departments';
var WORKLOCS_SHEET_NAME = 'WorkLocations';
var PRIMARYCAUSE_SHEET_NAME = 'PrimaryCauses';
var TREATMENTS_SHEET_NAME = 'Treatments';
var PROVIDERS_SHEET_NAME = 'TreatmentProviders';
var BODYPARTS_SHEET_NAME = 'BodyParts';

// --- HTML INCLUDE HELPER ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- ROUTING ---
function doGet(e) {
  var page = 'index';
  if (e && e.parameter && e.parameter.page) page = e.parameter.page;

  // 1. Provider Directory (Public)
  if (page == 'providers') {
    return HtmlService.createTemplateFromFile('Providers').evaluate()
      .setTitle('Treatment Providers')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 2. Signed Notice View (Public but secured by signature)
  if (page == 'notice') {
    var t = HtmlService.createTemplateFromFile('Notice');
    t.rid = ''; t.row = ''; t.ts = ''; t.sig = '';
    t.noticeError = ''; t.noticeJson = 'null';

    if (e && e.parameter) {
      if (e.parameter.rid) t.rid = e.parameter.rid;
      if (e.parameter.row) t.row = e.parameter.row;
      if (e.parameter.ts) t.ts = e.parameter.ts;
      if (e.parameter.sig) t.sig = e.parameter.sig;
    }

    var v = verifyNoticeSig_(t.rid, t.row, t.ts, t.sig);
    if (!v.ok) {
      t.noticeError = v.reason;
    } else {
      try {
        // Function located in EmailService.gs (or CaseService.gs)
        var noticeObj = getNoticeData(t.rid, t.row);
        t.noticeJson = JSON.stringify(noticeObj);
      } catch (err) {
        t.noticeError = 'invalid';
      }
    }

    return t.evaluate()
      .setTitle('Injury Update Notice')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 3. Admin Reports Dashboard
  if (page == 'reports') {
    return HtmlService.createTemplateFromFile('Reports').evaluate()
      .setTitle('FROI Admin Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 4. Default: Index (FROI Form)
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('City of Portland FROI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// --- SHARED UTILS ---
function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  } else {
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    // If row 1 is blank, set headers
    var lastCol = Math.max(sh.getLastColumn(), headers.length);
    var r1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var empty = r1.every(function (v) { return String(v).trim() === ''; });
    if (empty) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isAllowedDomain_(email) {
  var em = normalizeEmail_(email);
  return em && em.indexOf(ALLOWED_DOMAIN) !== -1;
}

function parseBool_(v) {
  var s = String(v || '').trim().toLowerCase();
  return (s === 'y' || s === 'yes' || s === 'true' || s === '1');
}

function safe_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

/**
 * EXISTING FUNCTION (keep it for backwards compatibility)
 */
function getCurrentUserEmail_() {
  try {
    var e = Session.getActiveUser().getEmail();
    if (e) return e;
  } catch (err1) {}

  try {
    var e2 = Session.getEffectiveUser().getEmail();
    if (e2) return e2;
  } catch (err2) {}

  return '';
}

/**
 * FINAL FIX: PUBLIC ALIAS WITHOUT UNDERSCORE
 * This is what your client should call via google.script.run.getCurrentUserEmail()
 */
function getCurrentUserEmail() {
  return getCurrentUserEmail_();
}
// ===============================
// SETTINGS BACKEND (Modal_Settings)
// ===============================

function _ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _tz_() {
  return Session.getScriptTimeZone() || 'America/New_York';
}

function _now_() {
  return Utilities.formatDate(new Date(), _tz_(), 'yyyy-MM-dd HH:mm');
}

function _sheet_(name, headers) {
  return ensureSheet_(name, headers);
}

function _headerMap_(headersRow) {
  var m = {};
  (headersRow || []).forEach(function(h, i) {
    var k = String(h || '').trim();
    if (k) m[k] = i;
  });
  return m;
}

function _col_(map, candidates, fallbackIndex) {
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    if (map.hasOwnProperty(c)) return map[c];
  }
  return fallbackIndex;
}

function _requireAdmin_(email) {
  var em = normalizeEmail_(email);
  if (!em) return { ok: false, error: 'Unable to determine your email for admin validation.' };
  if (!isAllowedDomain_(em)) return { ok: false, error: 'Access denied (domain).' };

  var sh = _sheet_(ADMIN_SHEET_NAME, ['Email', 'Name', 'Title', 'Department']);
  var vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return { ok: false, error: 'No admins configured.' };

  var hdr = _headerMap_(vals[0]);
  var cEmail = _col_(hdr, ['Email', 'email'], 0);

  for (var r = 1; r < vals.length; r++) {
    var rowEmail = normalizeEmail_(vals[r][cEmail]);
    if (rowEmail && rowEmail === em) return { ok: true };
  }
  return { ok: false, error: 'Access denied.' };
}

function _rows_(sh) {
  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return { headers: [], data: [] };
  return { headers: vals[0], data: vals.slice(1) };
}

// ---------- ADMINS ----------
function listAdmins(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(ADMIN_SHEET_NAME, ['Email', 'Name', 'Title', 'Department']);
  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  var cEmail = _col_(hdr, ['Email', 'email'], 0);
  var cName  = _col_(hdr, ['Name', 'name'], 1);
  var cTitle = _col_(hdr, ['Title', 'title'], 2);
  var cDept  = _col_(hdr, ['Department', 'department'], 3);

  var out = [];
  obj.data.forEach(function(r, i) {
    var em = safe_(r[cEmail]);
    if (!em) return;
    out.push({
      rowIdx: i + 2, // actual sheet row
      email: em,
      name: safe_(r[cName]),
      title: safe_(r[cTitle]),
      department: safe_(r[cDept])
    });
  });

  return { success: true, rows: out };
}

function addAdmin(currentEmail, payload) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var email = normalizeEmail_(payload && payload.email);
  if (!email) return { success: false, error: 'Email is required.' };
  if (!isAllowedDomain_(email)) return { success: false, error: 'Email must be a City address.' };

  var sh = _sheet_(ADMIN_SHEET_NAME, ['Email', 'Name', 'Title', 'Department']);
  sh.appendRow([
    email,
    safe_(payload.name),
    safe_(payload.title),
    safe_(payload.department)
  ]);
  return { success: true };
}

function updateAdmin(currentEmail, rowIdx, payload) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(ADMIN_SHEET_NAME, ['Email', 'Name', 'Title', 'Department']);
  var vals = sh.getDataRange().getValues();
  var hdr = _headerMap_(vals[0]);

  var cEmail = _col_(hdr, ['Email', 'email'], 0) + 1;
  var cName  = _col_(hdr, ['Name', 'name'], 1) + 1;
  var cTitle = _col_(hdr, ['Title', 'title'], 2) + 1;
  var cDept  = _col_(hdr, ['Department', 'department'], 3) + 1;

  var email = normalizeEmail_(payload && payload.email);
  if (!email) return { success: false, error: 'Email is required.' };

  sh.getRange(rowIdx, cEmail).setValue(email);
  sh.getRange(rowIdx, cName).setValue(safe_(payload.name));
  sh.getRange(rowIdx, cTitle).setValue(safe_(payload.title));
  sh.getRange(rowIdx, cDept).setValue(safe_(payload.department));
  return { success: true };
}

function deleteAdmin(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };
  var sh = _sheet_(ADMIN_SHEET_NAME, ['Email', 'Name', 'Title', 'Department']);
  sh.deleteRow(rowIdx);
  return { success: true };
}

// ---------- DEPARTMENTS ----------
function listDepartments(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(DEPTS_SHEET_NAME, [
    'Department','HR Name','HR Email','HR Phone',
    'Payroll Name','Payroll Email','Payroll Phone',
    'Safety Name','Safety Email','Safety Phone',
    'Updated'
  ]);

  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  var cDept = _col_(hdr, ['Department'], 0);
  var cHrN  = _col_(hdr, ['HR Name'], 1);
  var cHrE  = _col_(hdr, ['HR Email'], 2);
  var cHrP  = _col_(hdr, ['HR Phone'], 3);
  var cPrN  = _col_(hdr, ['Payroll Name'], 4);
  var cPrE  = _col_(hdr, ['Payroll Email'], 5);
  var cPrP  = _col_(hdr, ['Payroll Phone'], 6);
  var cSaN  = _col_(hdr, ['Safety Name'], 7);
  var cSaE  = _col_(hdr, ['Safety Email'], 8);
  var cSaP  = _col_(hdr, ['Safety Phone'], 9);
  var cUpd  = _col_(hdr, ['Updated'], 10);

  var out = [];
  obj.data.forEach(function(r, i) {
    var dept = safe_(r[cDept]);
    if (!dept) return;
    out.push({
      rowIdx: i + 2,
      department: dept,
      hrName: safe_(r[cHrN]),
      hrEmail: safe_(r[cHrE]),
      hrPhone: safe_(r[cHrP]),
      payrollName: safe_(r[cPrN]),
      payrollEmail: safe_(r[cPrE]),
      payrollPhone: safe_(r[cPrP]),
      safetyName: safe_(r[cSaN]),
      safetyEmail: safe_(r[cSaE]),
      safetyPhone: safe_(r[cSaP]),
      updated: safe_(r[cUpd])
    });
  });

  return { success: true, rows: out };
}

function addDepartment(currentEmail, payload) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var dept = safe_(payload && payload.department);
  if (!dept) return { success: false, error: 'Department name is required.' };

  var sh = _sheet_(DEPTS_SHEET_NAME, [
    'Department','HR Name','HR Email','HR Phone',
    'Payroll Name','Payroll Email','Payroll Phone',
    'Safety Name','Safety Email','Safety Phone',
    'Updated'
  ]);

  sh.appendRow([
    dept,
    safe_(payload.hrName), safe_(payload.hrEmail), safe_(payload.hrPhone),
    safe_(payload.payrollName), safe_(payload.payrollEmail), safe_(payload.payrollPhone),
    safe_(payload.safetyName), safe_(payload.safetyEmail), safe_(payload.safetyPhone),
    _now_()
  ]);

  return { success: true };
}

function updateDepartment(currentEmail, rowIdx, payload) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(DEPTS_SHEET_NAME, [
    'Department','HR Name','HR Email','HR Phone',
    'Payroll Name','Payroll Email','Payroll Phone',
    'Safety Name','Safety Email','Safety Phone',
    'Updated'
  ]);

  var vals = sh.getDataRange().getValues();
  var hdr = _headerMap_(vals[0]);

  function setByName(colName, value) {
    var idx = _col_(hdr, [colName], null);
    if (idx === null) return;
    sh.getRange(rowIdx, idx + 1).setValue(value);
  }

  var dept = safe_(payload && payload.department);
  if (!dept) return { success: false, error: 'Department name is required.' };

  setByName('Department', dept);
  setByName('HR Name', safe_(payload.hrName));
  setByName('HR Email', safe_(payload.hrEmail));
  setByName('HR Phone', safe_(payload.hrPhone));
  setByName('Payroll Name', safe_(payload.payrollName));
  setByName('Payroll Email', safe_(payload.payrollEmail));
  setByName('Payroll Phone', safe_(payload.payrollPhone));
  setByName('Safety Name', safe_(payload.safetyName));
  setByName('Safety Email', safe_(payload.safetyEmail));
  setByName('Safety Phone', safe_(payload.safetyPhone));
  setByName('Updated', _now_());

  return { success: true };
}

function deleteDepartment(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(DEPTS_SHEET_NAME, [
    'Department','HR Name','HR Email','HR Phone',
    'Payroll Name','Payroll Email','Payroll Phone',
    'Safety Name','Safety Email','Safety Phone',
    'Updated'
  ]);
  sh.deleteRow(rowIdx);
  return { success: true };
}

// ---------- WORK LOCATIONS ----------
function listWorkLocations(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(WORKLOCS_SHEET_NAME, ['Work Location','Department','Updated']);
  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  var cWl  = _col_(hdr, ['Work Location'], 0);
  var cDep = _col_(hdr, ['Department'], 1);
  var cUpd = _col_(hdr, ['Updated'], 2);

  var out = [];
  obj.data.forEach(function(r, i) {
    var wl = safe_(r[cWl]);
    if (!wl) return;
    out.push({
      rowIdx: i + 2,
      workLocation: wl,
      department: safe_(r[cDep]),
      updated: safe_(r[cUpd])
    });
  });

  return { success: true, rows: out };
}

function addWorkLocation(currentEmail, name, dept) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  name = safe_(name); dept = safe_(dept);
  if (!name) return { success: false, error: 'Work Location is required.' };
  if (!dept) return { success: false, error: 'Department is required.' };

  var sh = _sheet_(WORKLOCS_SHEET_NAME, ['Work Location','Department','Updated']);
  sh.appendRow([name, dept, _now_()]);
  return { success: true };
}

function updateWorkLocation(currentEmail, rowIdx, name, dept) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  name = safe_(name); dept = safe_(dept);
  if (!name) return { success: false, error: 'Work Location is required.' };
  if (!dept) return { success: false, error: 'Department is required.' };

  var sh = _sheet_(WORKLOCS_SHEET_NAME, ['Work Location','Department','Updated']);
  sh.getRange(rowIdx, 1).setValue(name);
  sh.getRange(rowIdx, 2).setValue(dept);
  sh.getRange(rowIdx, 3).setValue(_now_());
  return { success: true };
}

function deleteWorkLocation(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(WORKLOCS_SHEET_NAME, ['Work Location','Department','Updated']);
  sh.deleteRow(rowIdx);
  return { success: true };
}

// ---------- PRIMARY CAUSES ----------
function listPrimaryCauses(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(PRIMARYCAUSE_SHEET_NAME, ['Primary Cause','Updated']);
  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  var cCause = _col_(hdr, ['Primary Cause'], 0);
  var cUpd   = _col_(hdr, ['Updated'], 1);

  var out = [];
  obj.data.forEach(function(r, i) {
    var c = safe_(r[cCause]);
    if (!c) return;
    out.push({ rowIdx: i + 2, cause: c, updated: safe_(r[cUpd]) });
  });

  return { success: true, rows: out };
}

function addPrimaryCause(currentEmail, cause) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  cause = safe_(cause);
  if (!cause) return { success: false, error: 'Primary Cause is required.' };

  var sh = _sheet_(PRIMARYCAUSE_SHEET_NAME, ['Primary Cause','Updated']);
  sh.appendRow([cause, _now_()]);
  return { success: true };
}

function updatePrimaryCause(currentEmail, rowIdx, cause) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };
  cause = safe_(cause);
  if (!cause) return { success: false, error: 'Primary Cause is required.' };

  var sh = _sheet_(PRIMARYCAUSE_SHEET_NAME, ['Primary Cause','Updated']);
  sh.getRange(rowIdx, 1).setValue(cause);
  sh.getRange(rowIdx, 2).setValue(_now_());
  return { success: true };
}

function deletePrimaryCause(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(PRIMARYCAUSE_SHEET_NAME, ['Primary Cause','Updated']);
  sh.deleteRow(rowIdx);
  return { success: true };
}

// ---------- TREATMENTS ----------
function listTreatments(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(TREATMENTS_SHEET_NAME, ['Treatment','Updated']);
  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  var cT = _col_(hdr, ['Treatment'], 0);
  var cU = _col_(hdr, ['Updated'], 1);

  var out = [];
  obj.data.forEach(function(r, i) {
    var t = safe_(r[cT]);
    if (!t) return;
    out.push({ rowIdx: i + 2, treatment: t, updated: safe_(r[cU]) });
  });

  return { success: true, rows: out };
}

function addTreatment(currentEmail, treatment) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  treatment = safe_(treatment);
  if (!treatment) return { success: false, error: 'Treatment is required.' };

  var sh = _sheet_(TREATMENTS_SHEET_NAME, ['Treatment','Updated']);
  sh.appendRow([treatment, _now_()]);
  return { success: true };
}

function updateTreatment(currentEmail, rowIdx, treatment) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };
  treatment = safe_(treatment);
  if (!treatment) return { success: false, error: 'Treatment is required.' };

  var sh = _sheet_(TREATMENTS_SHEET_NAME, ['Treatment','Updated']);
  sh.getRange(rowIdx, 1).setValue(treatment);
  sh.getRange(rowIdx, 2).setValue(_now_());
  return { success: true };
}

function deleteTreatment(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(TREATMENTS_SHEET_NAME, ['Treatment','Updated']);
  sh.deleteRow(rowIdx);
  return { success: true };
}

// ---------- PROVIDERS ----------
function listTreatmentProviders(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(PROVIDERS_SHEET_NAME, [
    'Provider Name','Address','Phone',
    'Mon Open','Mon Hours','Tue Open','Tue Hours','Wed Open','Wed Hours',
    'Thu Open','Thu Hours','Fri Open','Fri Hours','Sat Open','Sat Hours',
    'Sun Open','Sun Hours',
    'Updated'
  ]);

  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  function g(r, name, fallback) {
    var idx = _col_(hdr, [name], fallback);
    return r[idx];
  }

  var out = [];
  obj.data.forEach(function(r, i) {
    var name = safe_(g(r, 'Provider Name', 0));
    if (!name) return;

    out.push({
      rowIdx: i + 2,
      providerName: name,
      address: safe_(g(r, 'Address', 1)),
      phone: safe_(g(r, 'Phone', 2)),
      monOpen: parseBool_(g(r, 'Mon Open', 3)),
      monHours: safe_(g(r, 'Mon Hours', 4)),
      tueOpen: parseBool_(g(r, 'Tue Open', 5)),
      tueHours: safe_(g(r, 'Tue Hours', 6)),
      wedOpen: parseBool_(g(r, 'Wed Open', 7)),
      wedHours: safe_(g(r, 'Wed Hours', 8)),
      thuOpen: parseBool_(g(r, 'Thu Open', 9)),
      thuHours: safe_(g(r, 'Thu Hours', 10)),
      friOpen: parseBool_(g(r, 'Fri Open', 11)),
      friHours: safe_(g(r, 'Fri Hours', 12)),
      satOpen: parseBool_(g(r, 'Sat Open', 13)),
      satHours: safe_(g(r, 'Sat Hours', 14)),
      sunOpen: parseBool_(g(r, 'Sun Open', 15)),
      sunHours: safe_(g(r, 'Sun Hours', 16)),
      updated: safe_(g(r, 'Updated', 17))
    });
  });

  return { success: true, rows: out };
}

function addTreatmentProvider(currentEmail, payload) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  if (!payload || !safe_(payload.providerName)) return { success: false, error: 'Provider Name is required.' };

  var sh = _sheet_(PROVIDERS_SHEET_NAME, [
    'Provider Name','Address','Phone',
    'Mon Open','Mon Hours','Tue Open','Tue Hours','Wed Open','Wed Hours',
    'Thu Open','Thu Hours','Fri Open','Fri Hours','Sat Open','Sat Hours',
    'Sun Open','Sun Hours',
    'Updated'
  ]);

  sh.appendRow([
    safe_(payload.providerName),
    safe_(payload.address),
    safe_(payload.phone),
    payload.monOpen ? 'Y' : '',
    safe_(payload.monHours),
    payload.tueOpen ? 'Y' : '',
    safe_(payload.tueHours),
    payload.wedOpen ? 'Y' : '',
    safe_(payload.wedHours),
    payload.thuOpen ? 'Y' : '',
    safe_(payload.thuHours),
    payload.friOpen ? 'Y' : '',
    safe_(payload.friHours),
    payload.satOpen ? 'Y' : '',
    safe_(payload.satHours),
    payload.sunOpen ? 'Y' : '',
    safe_(payload.sunHours),
    _now_()
  ]);

  return { success: true };
}

function updateTreatmentProvider(currentEmail, rowIdx, payload) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };
  if (!payload || !safe_(payload.providerName)) return { success: false, error: 'Provider Name is required.' };

  var sh = _sheet_(PROVIDERS_SHEET_NAME, [
    'Provider Name','Address','Phone',
    'Mon Open','Mon Hours','Tue Open','Tue Hours','Wed Open','Wed Hours',
    'Thu Open','Thu Hours','Fri Open','Fri Hours','Sat Open','Sat Hours',
    'Sun Open','Sun Hours',
    'Updated'
  ]);

  var row = [
    safe_(payload.providerName),
    safe_(payload.address),
    safe_(payload.phone),
    payload.monOpen ? 'Y' : '',
    safe_(payload.monHours),
    payload.tueOpen ? 'Y' : '',
    safe_(payload.tueHours),
    payload.wedOpen ? 'Y' : '',
    safe_(payload.wedHours),
    payload.thuOpen ? 'Y' : '',
    safe_(payload.thuHours),
    payload.friOpen ? 'Y' : '',
    safe_(payload.friHours),
    payload.satOpen ? 'Y' : '',
    safe_(payload.satHours),
    payload.sunOpen ? 'Y' : '',
    safe_(payload.sunHours),
    _now_()
  ];

  sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  return { success: true };
}

function deleteTreatmentProvider(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(PROVIDERS_SHEET_NAME, [
    'Provider Name','Address','Phone',
    'Mon Open','Mon Hours','Tue Open','Tue Hours','Wed Open','Wed Hours',
    'Thu Open','Thu Hours','Fri Open','Fri Hours','Sat Open','Sat Hours',
    'Sun Open','Sun Hours',
    'Updated'
  ]);
  sh.deleteRow(rowIdx);
  return { success: true };
}

// ---------- BODY PARTS ----------
function listBodyParts(currentEmail) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };

  var sh = _sheet_(BODYPARTS_SHEET_NAME, ['Body Part','Updated']);
  var obj = _rows_(sh);
  var hdr = _headerMap_(obj.headers);

  var cB = _col_(hdr, ['Body Part'], 0);
  var cU = _col_(hdr, ['Updated'], 1);

  var out = [];
  obj.data.forEach(function(r, i) {
    var bp = safe_(r[cB]);
    if (!bp) return;
    out.push({ rowIdx: i + 2, bodyPart: bp, updated: safe_(r[cU]) });
  });

  return { success: true, rows: out };
}

function addBodyPart(currentEmail, bodyPart) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  bodyPart = safe_(bodyPart);
  if (!bodyPart) return { success: false, error: 'Body Part is required.' };

  var sh = _sheet_(BODYPARTS_SHEET_NAME, ['Body Part','Updated']);
  sh.appendRow([bodyPart, _now_()]);
  return { success: true };
}

function updateBodyPart(currentEmail, rowIdx, bodyPart) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  bodyPart = safe_(bodyPart);
  if (!bodyPart) return { success: false, error: 'Body Part is required.' };

  var sh = _sheet_(BODYPARTS_SHEET_NAME, ['Body Part','Updated']);
  sh.getRange(rowIdx, 1).setValue(bodyPart);
  sh.getRange(rowIdx, 2).setValue(_now_());
  return { success: true };
}

function deleteBodyPart(currentEmail, rowIdx) {
  var v = _requireAdmin_(currentEmail);
  if (!v.ok) return { success: false, error: v.error };
  rowIdx = Number(rowIdx);
  if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row.' };

  var sh = _sheet_(BODYPARTS_SHEET_NAME, ['Body Part','Updated']);
  sh.deleteRow(rowIdx);
  return { success: true };
}

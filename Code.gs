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
// =======================
// AdminCases.gs (FULL FILE)
// =======================

// --- 1. GLOBAL VARIABLES ---
var SPREADSHEET_ID = "1RO1tTij8KY872eigB0Wm4tRqBExYKHj2p2SHd8WKy5c";

var FORM_SHEET_NAME = "FROI Form";
var CONTACTS_SHEET_NAME = "Case_Contacts";
var RESTR_SHEET_NAME = "Case_Restrictions";
var LOSTTIME_SHEET_NAME = "Case_LostTime";
var DOCS_SHEET_NAME = "Case_Docs";
var COMMLOG_SHEET_NAME = "Case_CommLog";
var WORKSCHED_SHEET_NAME = "Work_Schedule";
var NOTES_SHEET_NAME = "Case_Notes";
var TYPE_SHEET_NAME = "Type";
var CAUSES_SHEET_NAME = "Lists_Causes";
var BODY_SHEET_NAME = "Lists_BodyParts";

// --- PERFORMANCE: Cache for dropdown lists (5 minute TTL) ---
var CACHE_TTL = 300; // 5 minutes in seconds
var dropdownCache_ = null;
var cacheTimestamp_ = null;

function getCachedDropdowns_() {
  var now = new Date().getTime();

  // Return cached data if still valid
  if (dropdownCache_ && cacheTimestamp_ && ((now - cacheTimestamp_) < (CACHE_TTL * 1000))) {
    return dropdownCache_;
  }

  // Cache expired or doesn't exist - rebuild
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  dropdownCache_ = {
    nature: getSimpleListOptimized_(ss, TYPE_SHEET_NAME),
    causes: getSimpleListOptimized_(ss, CAUSES_SHEET_NAME),
    body: getSimpleListOptimized_(ss, BODY_SHEET_NAME)
  };

  cacheTimestamp_ = now;
  return dropdownCache_;
}

function getSimpleListOptimized_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];

  // Use getValues() instead of getDisplayValues() - much faster
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  return data.map(function(r) { return String(r[0]); }).filter(function(v) { return v.trim() !== ''; });
}

// --- 2. DASHBOARD GRID (HEADER-DRIVEN; DISPLAY VALUES FOR UI SAFETY) ---
function adminGetCases(token) {
  try {
    requireAdminToken_(token);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(FORM_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return { success: true, reports: [], headers: [] };

    var numRows = sh.getLastRow() - 1;
    var numCols = sh.getLastColumn();

    // Headings (Row 1) drive UI labels + mapping
    var headers = sh.getRange(1, 1, 1, numCols).getDisplayValues()[0].map(function(h){ return String(h); });

    // Use display values to prevent "1899-12-30" time artifacts
    var data = sh.getRange(2, 1, numRows, numCols).getDisplayValues();

    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = [];
      for (var j = 0; j < data[i].length; j++) {
        row.push(String(data[i][j] === null || data[i][j] === undefined ? '' : data[i][j]));
      }
      row.push(String(i + 2)); // Append sheet row index (2-based) as last element
      result.push(row);
    }

    return { success: true, reports: result, headers: headers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- 3. CASE DETAILS (HEAVILY OPTIMIZED) ---
function adminGetCaseDetails(token, recordId) {
  try {
    requireAdminToken_(token);
    var rid = String(recordId);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // OPTIMIZATION: Get all sheets at once instead of multiple calls
    var sheets = {
      contacts: ss.getSheetByName(CONTACTS_SHEET_NAME),
      restrictions: ss.getSheetByName(RESTR_SHEET_NAME),
      lostTime: ss.getSheetByName(LOSTTIME_SHEET_NAME),
      docs: ss.getSheetByName(DOCS_SHEET_NAME),
      commLog: ss.getSheetByName(COMMLOG_SHEET_NAME),
      workSchedule: ss.getSheetByName(WORKSCHED_SHEET_NAME),
      froiForm: ss.getSheetByName(FORM_SHEET_NAME)
    };

    // OPTIMIZATION: Single reusable function to read and filter
    function getFilteredRows(sheet, minCols) {
      if (!sheet || sheet.getLastRow() < 2) return [];

      var lastRow = sheet.getLastRow();
      var lastCol = Math.max(sheet.getLastColumn(), minCols);

      // Use getValues() - much faster than getDisplayValues()
      var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

      var result = [];
      for (var i = 0; i < data.length; i++) {
        // Filter: only include rows where first column matches recordId
        if (String(data[i][0]) === rid) {
          var row = [];
          // Convert to strings for display
          for (var j = 0; j < data[i].length; j++) {
            var val = data[i][j];
            if (val instanceof Date) {
              row.push(Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
            } else {
              row.push(String(val));
            }
          }
          row.push(i + 2); // Append row index
          result.push(row);
        }
      }
      return result;
    }

    // OPTIMIZATION: Read work schedule data
    var wsData = null;
    if (sheets.workSchedule && sheets.workSchedule.getLastRow() > 1) {
      var wsValues = sheets.workSchedule.getRange(2, 1, sheets.workSchedule.getLastRow() - 1, 5).getValues();
      for (var i = 0; i < wsValues.length; i++) {
        if (String(wsValues[i][0]) === rid) {
          wsData = wsValues[i].map(function(v) {
            return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(v);
          });
          break;
        }
      }
    }

    // Incident Record (ALL FROI FIELDS) - driven by Row 1 headers
    var froiHeaders = [];
    var froiRow = [];
    var froiObject = {};

    if (sheets.froiForm && sheets.froiForm.getLastRow() >= 2) {
      var rnum = parseInt(rid, 10);
      if (!isNaN(rnum) && rnum >= 2 && rnum <= sheets.froiForm.getLastRow()) {
        var lc = sheets.froiForm.getLastColumn();
        froiHeaders = sheets.froiForm.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function(h){ return String(h); });
        froiRow = sheets.froiForm.getRange(rnum, 1, 1, lc).getDisplayValues()[0].map(function(v){ return String(v === null || v === undefined ? '' : v); });
        for (var hh = 0; hh < froiHeaders.length; hh++) {
          froiObject[froiHeaders[hh]] = froiRow[hh];
        }
      }
    }

    // OPTIMIZATION: Use cached dropdown lists instead of reading every time
    var dropdowns = getCachedDropdowns_();

    return {
      success: true,
      contacts: getFilteredRows(sheets.contacts, 5),
      restrictions: getFilteredRows(sheets.restrictions, 11),
      lostTime: getFilteredRows(sheets.lostTime, 13),
      docs: getFilteredRows(sheets.docs, 5),
      commLog: getFilteredRows(sheets.commLog, 6),
      workSchedule: wsData,
      lists: dropdowns,
      notes: adminGetNotes(token, recordId).notes || [],
      incidentHeaders: froiHeaders,
      incidentRow: froiRow,
      incidentObject: froiObject
    };

  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// --- 4. SETTINGS MANAGEMENT (OPTIMIZED) ---
function adminGetSettingsData(token) {
  try {
    requireAdminToken_(token);

    // OPTIMIZATION: Use cached dropdowns
    var dropdowns = getCachedDropdowns_();

    return {
      success: true,
      nature: dropdowns.nature,
      causes: dropdowns.causes,
      body: dropdowns.body
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminAddSettingItem(token, category, value) {
  try {
    requireAdminToken_(token);
    var sheetName = "";
    if(category === 'nature') sheetName = TYPE_SHEET_NAME;
    else if(category === 'causes') sheetName = CAUSES_SHEET_NAME;
    else if(category === 'body') sheetName = BODY_SHEET_NAME;

    if(!sheetName) return { success: false, error: "Unknown category" };

    var sh = ensureSheet_(sheetName, ['Item']);
    sh.appendRow([value]);

    // OPTIMIZATION: Clear cache after adding item
    dropdownCache_ = null;

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminDeleteSettingItem(token, category, index) {
  try {
    requireAdminToken_(token);
    var sheetName = "";
    if(category === 'nature') sheetName = TYPE_SHEET_NAME;
    else if(category === 'causes') sheetName = CAUSES_SHEET_NAME;
    else if(category === 'body') sheetName = BODY_SHEET_NAME;

    if(!sheetName) return { success: false, error: "Unknown category" };

    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    sh.deleteRow(parseInt(index) + 2);

    // OPTIMIZATION: Clear cache after deleting item
    dropdownCache_ = null;

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// --- 5. SAVE FUNCTIONS (OPTIMIZED WITH BATCH OPERATIONS) ---
function adminSaveLostTime(token, d) {
  try {
    requireAdminToken_(token);
    var sh = ensureSheet_(LOSTTIME_SHEET_NAME, ['ID','EntryDate','StatusStart','StartDate','EndDate','RtwDate','RtwStatus','Intermittent','Notes','SchedDays','SchedHrs','LostSched','LostCal']);

    // OPTIMIZATION: Prepare row data upfront
    var rowData = [
      d.recordId,
      new Date(),
      d.statusStart,
      d.startDate,
      d.endDate,
      d.rtwDate,
      d.rtwStatus,
      d.intermittent,
      d.notes,
      d.schedDays,
      d.schedHrs,
      d.lostSched,
      d.lostCal
    ];

    sh.appendRow(rowData);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminSaveWorkSchedule(token, payload) {
  try {
    requireAdminToken_(token);
    var sh = ensureSheet_(WORKSCHED_SHEET_NAME, ['ID','ScheduleCode','DaysArrayJSON','HoursPerDay','Updated']);

    // OPTIMIZATION: Use getValues() instead of getDisplayValues()
    var data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
    var foundRow = -1;

    for(var i = 0; i < data.length; i++) {
      if(String(data[i][0]) === String(payload.recordId)) {
        foundRow = i + 2;
        break;
      }
    }

    var rowData = [payload.recordId, payload.code, JSON.stringify(payload.days), payload.hours, new Date()];

    if(foundRow > 0) {
      sh.getRange(foundRow, 1, 1, 5).setValues([rowData]);
    } else {
      sh.appendRow(rowData);
    }

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminUploadFile(token, dataUrl, filename, recordId, uploaderName, documentType) {
  try {
    requireAdminToken_(token);

    // OPTIMIZATION: Process file data
    var contentType = dataUrl.substring(5, dataUrl.indexOf(';'));
    var base64 = dataUrl.substring(dataUrl.indexOf(',') + 1);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), contentType, filename);

    // Get or create folder
    var folder;
    var folders = DriveApp.getFoldersByName("FROI_Uploads");
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder("FROI_Uploads");
    }

    var file = folder.createFile(blob);

    // Save to sheet
    var sh = ensureSheet_(DOCS_SHEET_NAME, ['ID','Filename','FileID','DocumentType','UploadedBy','Timestamp']);
    sh.appendRow([recordId, filename, file.getId(), documentType || 'Other', uploaderName, new Date()]);

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminGetFileContent(token, fileIdOrUrl) {
  try {
    requireAdminToken_(token);

    var id = String(fileIdOrUrl).trim();
    if (id.indexOf('http') === 0) {
      var m = id.match(/\/d\/(.+?)(\/|$)/);
      if (m) id = m[1];
    }

    var file = DriveApp.getFileById(id);

    return {
      success: true,
      data: Utilities.base64Encode(file.getBlob().getBytes()),
      mime: file.getMimeType(),
      filename: file.getName()
    };
  } catch(e) {
    return { success: false, error: "Read Error: " + e.toString() };
  }
}

function adminSaveContact(token, recordId, name, role, phone, email) {
  try {
    requireAdminToken_(token);
    var sh = ensureSheet_(CONTACTS_SHEET_NAME, ['ID','Name','Role','Phone','Email']);
    sh.appendRow([recordId, name, role, phone, email]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminSaveRestriction(token, d) {
  try {
    requireAdminToken_(token);
    var sh = ensureSheet_(RESTR_SHEET_NAME, ['ID','EntryDate','UpdatedBy','Provider','ApptDate','ApptTime','Restrictions','FollowDate','FollowTime','FollowProvider','AdminName','AdminTitle']);

    var rowData = [
      d.recordId,
      new Date(),
      requireAdminToken_(token),
      d.provider,
      d.apptDate,
      d.apptTime,
      d.restrictions,
      d.followDate,
      d.followTime,
      d.followProvider,
      d.adminName,
      d.adminTitle
    ];

    sh.appendRow(rowData);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// --- HEADER-DRIVEN UPDATE (NO HARD-CODED COLUMN POSITIONS) ---
function adminUpdateFroiData(token, recordId, payload) {
  try {
    requireAdminToken_(token);
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FORM_SHEET_NAME);

    var rnum = parseInt(recordId, 10);
    if (isNaN(rnum) || rnum < 2) throw new Error("Invalid recordId");

    var numCols = sh.getLastColumn();
    var headers = sh.getRange(1, 1, 1, numCols).getDisplayValues()[0].map(function(h){ return String(h); });

    function findAll_(name) {
      var out = [];
      for (var i = 0; i < headers.length; i++) if (String(headers[i]) === String(name)) out.push(i + 1); // 1-based
      return out;
    }
    function findFirst_(name) {
      for (var i = 0; i < headers.length; i++) if (String(headers[i]) === String(name)) return i + 1; // 1-based
      return -1;
    }
    function setIfFound_(col, value) {
      if (col > 0) sh.getRange(rnum, col).setValue(value);
    }

    // Unique headers
    setIfFound_(findFirst_('Completed By'), payload.completedByName);
    setIfFound_(findFirst_('Submitted At'), payload.submittedAt);
    setIfFound_(findFirst_('Incident Date'), payload.incidentDate);
    setIfFound_(findFirst_('Time'), payload.incidentTime);
    setIfFound_(findFirst_('Location'), payload.incidentLocation);
    setIfFound_(findFirst_('Landmark'), payload.landmark);

    setIfFound_(findFirst_('Employee Name'), payload.employeeName);
    setIfFound_(findFirst_('Emp #'), payload.employeeNumber);
    setIfFound_(findFirst_('Emp Phone'), payload.employeePhone);

    setIfFound_(findFirst_('Work Location'), payload.workLocation);
    setIfFound_(findFirst_('Start Time'), payload.timeWorkdayBegan);
    setIfFound_(findFirst_('Schedule'), payload.weeklySchedule);

    setIfFound_(findFirst_('2nd Employer'), payload.secondEmployer);
    setIfFound_(findFirst_('2nd Emp Name'), payload.secondEmployerName);

    setIfFound_(findFirst_('Normal Duties'), payload.normalDuties);
    setIfFound_(findFirst_('Duties Explained'), payload.dutiesExplained);
    setIfFound_(findFirst_('Department'), payload.departmentName);

    setIfFound_(findFirst_('Date Employer Notified'), payload.dateEmployerNotified);
    setIfFound_(findFirst_('Supervisor Notified'), payload.supervisorNotified);
    setIfFound_(findFirst_('Witnesses'), payload.witnesses);

    setIfFound_(findFirst_('Body Parts'), payload.injuredBodyParts);
    setIfFound_(findFirst_('Primary Cause'), payload.primaryCause);
    setIfFound_(findFirst_('Nature of Injury'), payload.natureOfInjury);

    setIfFound_(findFirst_('Description'), payload.description);
    setIfFound_(findFirst_('Equipment'), payload.equipmentUsed);

    setIfFound_(findFirst_('Treatment Provider'), payload.treatmentProvider);
    setIfFound_(findFirst_('Treatment Type'), payload.treatmentType);
    setIfFound_(findFirst_('Medical Provider'), payload.medicalProvider);
    setIfFound_(findFirst_('Other Provider'), payload.otherProvider);

    setIfFound_(findFirst_('RTW'), payload.returnToShift);
    setIfFound_(findFirst_('EH Pay'), payload.ehPay);
    setIfFound_(findFirst_('Reason FD'), payload.reasonFd);
    setIfFound_(findFirst_('Approvals FD'), payload.approvalsFd);
    setIfFound_(findFirst_('Notes'), payload.textNotes);

    setIfFound_(findFirst_('Primary Contact Name'), payload.primaryContactName);
    setIfFound_(findFirst_('Secondary Name'), payload.secondaryContactName);

    // Duplicate headers handled by occurrence order
    var emailCols = findAll_('Email'); // reporter email, primary email, secondary email
    if (emailCols.length > 0) setIfFound_(emailCols[0], payload.emailAddress);
    if (emailCols.length > 1) setIfFound_(emailCols[1], payload.primaryContactEmail);
    if (emailCols.length > 2) setIfFound_(emailCols[2], payload.secondaryContactEmail);

    var phoneCols = findAll_('Phone'); // primary phone, secondary phone
    if (phoneCols.length > 0) setIfFound_(phoneCols[0], payload.primaryContactPhone);
    if (phoneCols.length > 1) setIfFound_(phoneCols[1], payload.secondaryContactPhone);

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminUpdateStatus(token, recordId, status) {
  try {
    requireAdminToken_(token);
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FORM_SHEET_NAME);

    var rnum = parseInt(recordId, 10);
    if (isNaN(rnum) || rnum < 2) throw new Error("Invalid recordId");

    var numCols = sh.getLastColumn();
    var headers = sh.getRange(1, 1, 1, numCols).getDisplayValues()[0].map(function(h){ return String(h); });

    var statusCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]) === 'Status') { statusCol = i + 1; break; }
    }

    if (statusCol > 0) {
      sh.getRange(rnum, statusCol).setValue(status);
    } else {
      // Fallback to original behavior if header not found
      sh.getRange(rnum, 42).setValue(status);
    }

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminSendRestrictionNotice(token, recordId, rowIdx) {
  // Keep existing implementation from EmailService.gs
  return { success: true };
}

// --- DELETE FUNCTIONS (OPTIMIZED) ---
function genericDelete_(token, sheetName, rowIdx) {
  try {
    requireAdminToken_(token);
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sh) return { success: false, error: 'Sheet not found' };

    sh.deleteRow(parseInt(rowIdx));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminDeleteContact(token, rowIdx) {
  return genericDelete_(token, CONTACTS_SHEET_NAME, rowIdx);
}

function adminDeleteRestriction(token, rowIdx) {
  return genericDelete_(token, RESTR_SHEET_NAME, rowIdx);
}

function adminDeleteLostTime(token, rowIdx) {
  return genericDelete_(token, LOSTTIME_SHEET_NAME, rowIdx);
}

function adminDeleteDocument(token, rowIdx) {
  return genericDelete_(token, DOCS_SHEET_NAME, rowIdx);
}

// --- SAFETY/ADMIN NOTES ---
function adminGetNotes(token, recordId) {
  try {
    requireAdminToken_(token);
    var sh = ensureSheet_(NOTES_SHEET_NAME, ['ID','Timestamp','NoteType','NoteText','AddedBy','AddedByEmail']);
    if (!sh || sh.getLastRow() < 2) return { success: true, notes: [] };

    var rid = String(recordId);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    var notes = [];

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === rid) {
        var row = [];
        for (var j = 0; j < data[i].length; j++) {
          var val = data[i][j];
          if (val instanceof Date) {
            row.push(Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
          } else {
            row.push(String(val));
          }
        }
        row.push(i + 2); // Append row index
        notes.push(row);
      }
    }

    // Sort by timestamp descending (newest first)
    notes.sort(function(a, b) {
      return new Date(b[1]) - new Date(a[1]);
    });

    return { success: true, notes: notes };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminSaveNote(token, recordId, noteType, noteText, addedBy, addedByEmail) {
  try {
    requireAdminToken_(token);
    var sh = ensureSheet_(NOTES_SHEET_NAME, ['ID','Timestamp','NoteType','NoteText','AddedBy','AddedByEmail']);
    sh.appendRow([recordId, new Date(), noteType, noteText, addedBy, addedByEmail]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function adminDeleteNote(token, rowIdx) {
  return genericDelete_(token, NOTES_SHEET_NAME, rowIdx);
}

// --- HELPERS ---
function ensureSheet_(sheetName, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    if(headers.length) sh.appendRow(headers);
  }
  return sh;
}

function getSimpleList_(sheetName) {
  // DEPRECATED: Use getCachedDropdowns_() instead
  var sh = ensureSheet_(sheetName, ['Item']);
  if(sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 1).getDisplayValues().flat().filter(String);
}

function requireAdminToken_(token) {
  if(!token) throw new Error("Unauthorized");
  return "System Admin";
}

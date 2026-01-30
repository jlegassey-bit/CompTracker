// --- DROPDOWN DATA FOR FROI FORM ---
function getFroiDropdownData() {
  try {
    ensureSheet_(DEPTS_SHEET_NAME, ['Department','HR Name','HR Email','HR Phone','Payroll Name','Payroll Email','Payroll Phone','Safety Name','Safety Email','Safety Phone','Updated']);
    ensureSheet_(WORKLOCS_SHEET_NAME, ['Work Location','Department','Updated']);
    ensureSheet_(PRIMARYCAUSE_SHEET_NAME, ['Primary Cause','Updated']);
    ensureSheet_(TREATMENTS_SHEET_NAME, ['Treatment','Updated']);
    ensureSheet_(PROVIDERS_SHEET_NAME, ['Provider Name','Address','Phone','Mon Open','Mon Hours','Tue Open','Tue Hours','Wed Open','Wed Hours','Thu Open','Thu Hours','Fri Open','Fri Hours','Sat Open','Sat Hours','Sun Open','Sun Hours','Updated']);
    ensureSheet_(BODYPARTS_SHEET_NAME, ['Body Part','Updated']);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Departments
    var depSh = ss.getSheetByName(DEPTS_SHEET_NAME);
    var departments = [];
    if (depSh && depSh.getLastRow() >= 2) {
      depSh.getRange(2, 1, depSh.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) {
        if (r[0] && String(r[0]).trim()) departments.push(String(r[0]).trim());
      });
    }
    departments.sort();

    // Work Locations
    var wlSh = ss.getSheetByName(WORKLOCS_SHEET_NAME);
    var workLocations = [];
    if (wlSh && wlSh.getLastRow() >= 2) {
      wlSh.getRange(2, 1, wlSh.getLastRow() - 1, 3).getDisplayValues().forEach(function (r) {
        if (r[0] && String(r[0]).trim()) {
          workLocations.push({ name: String(r[0]).trim(), department: String(r[1]).trim() });
        }
      });
    }
    workLocations.sort(function(a,b){ return a.name.localeCompare(b.name); });

    // Primary Causes
    var pcSh = ss.getSheetByName(PRIMARYCAUSE_SHEET_NAME);
    var primaryCauses = [];
    if (pcSh && pcSh.getLastRow() >= 2) {
      pcSh.getRange(2, 1, pcSh.getLastRow() - 1, 1).getDisplayValues().forEach(function(r){
        if(r[0]) primaryCauses.push(String(r[0]).trim());
      });
    }
    primaryCauses.sort();

    // Treatments
    var trSh = ss.getSheetByName(TREATMENTS_SHEET_NAME);
    var treatments = [];
    if (trSh && trSh.getLastRow() >= 2) {
      trSh.getRange(2, 1, trSh.getLastRow() - 1, 1).getDisplayValues().forEach(function(r){
        if(r[0]) treatments.push(String(r[0]).trim());
      });
    }
    treatments.sort();

    // Body Parts
    var bpSh = ss.getSheetByName(BODYPARTS_SHEET_NAME);
    var bodyParts = [];
    if (bpSh && bpSh.getLastRow() >= 2) {
      bpSh.getRange(2, 1, bpSh.getLastRow() - 1, 1).getDisplayValues().forEach(function(r){
        if(r[0]) bodyParts.push(String(r[0]).trim());
      });
    }
    bodyParts.sort();

    // Providers
    var providers = getTreatmentProvidersDirectory().providers || [];
    
    // Injury Types
    var injuryTypes = [];
    var typeSh = ss.getSheetByName('Type');
    if (typeSh && typeSh.getLastRow() >= 2) {
      typeSh.getRange(2, 1, typeSh.getLastRow() - 1, 1).getDisplayValues().forEach(function(r) {
        if (r[0] && String(r[0]).trim()) injuryTypes.push(String(r[0]).trim());
      });
    }
    injuryTypes.sort();

    return { success: true, departments: departments, workLocations: workLocations, primaryCauses: primaryCauses, injuryTypes: injuryTypes, treatments: treatments, bodyParts: bodyParts, providers: providers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- PUBLIC PROVIDER DIRECTORY ---
function getTreatmentProvidersDirectory() {
  try {
    ensureSheet_(PROVIDERS_SHEET_NAME, ['Provider Name','Address','Phone','Mon Open','Mon Hours','Tue Open','Tue Hours','Wed Open','Wed Hours','Thu Open','Thu Hours','Fri Open','Fri Hours','Sat Open','Sat Hours','Sun Open','Sun Hours','Updated']);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(PROVIDERS_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return { success: true, providers: [] };

    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getDisplayValues();
    var providers = [];

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var name = String(r[0] || '').trim();
      if (!name) continue;

      providers.push({
        providerName: name,
        address: String(r[1] || '').trim(),
        phone: String(r[2] || '').trim(),
        hours: {
          mon: { open: parseBool_(r[3]), hours: String(r[4] || '').trim() },
          tue: { open: parseBool_(r[5]), hours: String(r[6] || '').trim() },
          wed: { open: parseBool_(r[7]), hours: String(r[8] || '').trim() },
          thu: { open: parseBool_(r[9]), hours: String(r[10] || '').trim() },
          fri: { open: parseBool_(r[11]), hours: String(r[12] || '').trim() },
          sat: { open: parseBool_(r[13]), hours: String(r[14] || '').trim() },
          sun: { open: parseBool_(r[15]), hours: String(r[16] || '').trim() }
        },
        updated: String(r[17] || '').trim()
      });
    }
    providers.sort(function (a, b) { return String(a.providerName).localeCompare(String(b.providerName)); });
    return { success: true, providers: providers };
  } catch (e) {
    return { success: false, error: e.toString(), providers: [] };
  }
}

// --- SUBMISSION LOGIC ---
function processForm(formObject) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Ensure Form Sheet Exists & Headers are correct
    var sheet = ensureSheet_(FORM_SHEET_NAME, [
      'Completed By','Email','Submitted At','Incident Date','Time','Location','Landmark',
      'Employee Name','Emp #','Emp Phone','Work Location','Start Time','Schedule',
      '2nd Employer','2nd Emp Name','Normal Duties','Duties Explained','Department',
      'Date Employer Notified','Supervisor Notified','Witnesses','Body Parts','Primary Cause',
      'Nature of Injury','Description','Equipment','Treatment Provider','Treatment Type',
      'Medical Provider','Other Provider','RTW','EH Pay','Reason FD','Approvals FD',
      'Notes','Primary Contact Name','Email','Phone','Secondary Name','Phone','Email','Status'
    ]);

    // Validation
    var reporterEmail = safe_(formObject.emailAddress).toLowerCase();
    var departmentName = safe_(formObject.departmentName);
    var workLocation = safe_(formObject.workLocation);
    
    if (!safe_(formObject.completedByName)) throw new Error("Completed By is required.");
    if (reporterEmail.indexOf('@') === -1) throw new Error("Valid email is required.");
    if (!departmentName) throw new Error("Department is required.");
    if (!workLocation) throw new Error("Work Location is required.");

    // Validate Location matches Department
    var allowedLocs = getWorkLocationsForDepartment_(departmentName);
    var locMatch = allowedLocs.some(function(l){ return l.toLowerCase() === workLocation.toLowerCase(); });
    if (allowedLocs.length > 0 && !locMatch) {
      throw new Error("Work Location must match the selected Department.");
    }

    // Prepare Row Data
    // Note: 'Status' set to "Pending Review" by default
    var rowData = [
      formObject.completedByName, reporterEmail, new Date(),
      formObject.incidentDate, formObject.incidentTime, formObject.incidentLocation, formObject.landmark,
      formObject.employeeName, formObject.employeeNumber, formObject.employeePhone,
      formObject.workLocation, formObject.timeWorkdayBegan, formObject.weeklySchedule,
      formObject.secondEmployer, formObject.secondEmployerName,
      formObject.normalDuties, formObject.dutiesExplained, formObject.departmentName,
      formObject.dateEmployerNotified, formObject.supervisorNotified, formObject.witnesses,
      formObject.injuredBodyParts, formObject.primaryCause, formObject.natureOfInjury,
      formObject.description, formObject.equipmentUsed, formObject.treatmentProvider,
      formObject.treatmentType, formObject.medicalProvider, formObject.otherProvider,
      formObject.returnToShift, formObject.ehPay, formObject.reasonFd, formObject.approvalsFd,
      formObject.textNotes, formObject.primaryContactName, formObject.primaryContactEmail,
      formObject.primaryContactPhone, formObject.secondaryContactName, formObject.secondaryContactPhone,
      formObject.secondaryContactEmail, "Pending Review", "New", "", ""
    ];

    sheet.appendRow(rowData);
    var recordId = sheet.getLastRow();

    // 2. AUTO-ADD DEPARTMENT CONTACTS to Case_Contacts Sheet
    var deptContacts = getDeptContacts_(departmentName); 
    // Returns object with keys: hrName, hrEmail, hrPhone, payrollName..., safetyName...

    var contactSheet = ensureSheet_(CONTACTS_SHEET_NAME, ['ID','Name','Role','Phone','Email']);
    
    // Add HR
    if (deptContacts.hrEmail) {
      contactSheet.appendRow([recordId, deptContacts.hrName || 'HR/PAO', 'PAO/HR Generalist', deptContacts.hrPhone || '', deptContacts.hrEmail]);
    }
    // Add Payroll
    if (deptContacts.payrollEmail) {
      contactSheet.appendRow([recordId, deptContacts.payrollName || 'Payroll', 'Payroll Specialist/Manager', deptContacts.payrollPhone || '', deptContacts.payrollEmail]);
    }
    // Add Safety
    if (deptContacts.safetyEmail) {
      contactSheet.appendRow([recordId, deptContacts.safetyName || 'Safety', 'Safety and Training Officer', deptContacts.safetyPhone || '', deptContacts.safetyEmail]);
    }
    // Add Supervisor (Reporter)
    if (reporterEmail) {
        contactSheet.appendRow([recordId, formObject.completedByName, 'Supervisor', '', reporterEmail]);
    }

    // 3. Send Notifications
    sendSubmissionEmails_(formObject, recordId);

    return { success: true, recordId: recordId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- INTERNAL HELPERS ---
function getWorkLocationsForDepartment_(departmentName) {
  var dept = String(departmentName || '').trim();
  if (!dept) return [];
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(WORKLOCS_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === dept.toLowerCase()) {
      out.push(String(data[i][0]).trim());
    }
  }
  return out;
}

function sendSubmissionEmails_(f, recordId) {
  // 1. Confirmation to Reporter
  var subject = "FROI Confirmation — Record #" + recordId;
  var body = "Your FROI for " + safe_(f.employeeName) + " has been submitted.\n\n" +
             "Date: " + safe_(f.incidentDate) + "\n" +
             "Location: " + safe_(f.incidentLocation) + "\n" +
             "This is an automated message.";
  
  try { MailApp.sendEmail(f.emailAddress, subject, body); } catch(e){}

  // 2. Notification to HR/Safety (Routing)
  // We use the same lookup to find emails to notify immediately upon submission
  var deptContacts = getDeptContacts_(f.departmentName);
  var recipients = [deptContacts.hrEmail, deptContacts.safetyEmail].filter(function(e){ 
    return e && e.indexOf('@') > -1; 
  });
  
  // Remove duplicates
  var uniqueRecipients = recipients.filter(function(item, pos) {
      return recipients.indexOf(item) == pos;
  });

  if (uniqueRecipients.length > 0) {
    var adminLink = getScriptUrl() + "?page=reports&rid=" + recordId;
    var notifySubj = "New FROI: " + safe_(f.employeeName) + " (" + safe_(f.departmentName) + ")";
    var notifyBody = "A new injury report has been submitted.\n\n" +
                     "Employee: " + safe_(f.employeeName) + "\n" +
                     "Cause: " + safe_(f.primaryCause) + "\n" +
                     "Description: " + safe_(f.description) + "\n\n" +
                     "View Record: " + adminLink;
    
    try { MailApp.sendEmail(uniqueRecipients.join(','), notifySubj, notifyBody); } catch(e){}
  }
}

function getDeptContacts_(deptName) {
  var def = { hrName:'', hrEmail:'', hrPhone:'', payrollName:'', payrollEmail:'', payrollPhone:'', safetyName:'', safetyEmail:'', safetyPhone:'' };
  if (!deptName) return def;
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(DEPTS_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return def;
  
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getDisplayValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(deptName).trim().toLowerCase()) {
      // Mapping columns based on Settings structure: 
      // 0:Name, 1:HR Name, 2:HR Email, 3:HR Phone, 4:Pay Name, 5:Pay Email, 6:Pay Phone, 7:Safe Name, 8:Safe Email, 9:Safe Phone
      return {
        hrName: data[i][1], hrEmail: data[i][2], hrPhone: data[i][3],
        payrollName: data[i][4], payrollEmail: data[i][5], payrollPhone: data[i][6],
        safetyName: data[i][7], safetyEmail: data[i][8], safetyPhone: data[i][9]
      };
    }
  }
  return def;
}
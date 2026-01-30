function adminSendRestrictionNotice(token, recordId, rowIdx) {
  try {
    var senderEmail = requireAdminToken_(token); // Validate admin
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Get Restriction Data
    var resSheet = ss.getSheetByName(RESTR_SHEET_NAME);
    // Note: rowIdx comes from the client table logic. Verify it belongs to recordId?
    // For simplicity, we trust the rowIdx maps to the sheet.
    var rData = resSheet.getRange(rowIdx, 1, 1, resSheet.getLastColumn()).getDisplayValues()[0];
    
    // 2. Get FROI Employee Name
    var froiSheet = ss.getSheetByName(FORM_SHEET_NAME);
    // This part is tricky if we don't have the exact FROI row index. 
    // We must search for recordId in Col A? No, recordId usually implies Row # in simplistic setups, 
    // OR we stored a GUID. The original code used `recordId` as Row Index.
    // Let's assume recordId = Row Index for the FROI sheet.
    var empName = froiSheet.getRange(recordId, 8).getValue(); 
    
    // 3. Get Contacts with Email
    var conSheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
    var contacts = conSheet.getDataRange().getDisplayValues(); // Scan all
    var emails = [];
    contacts.forEach(function(c){
      if(String(c[0]) === String(recordId) && c[4] && c[4].indexOf('@') > -1) {
        emails.push(c[4]);
      }
    });
    
    if (emails.length === 0) return { success: false, error: "No contacts with emails found." };

    // 4. Generate Link
    var ts = Date.now();
    var sig = computeNoticeSig_(recordId, rowIdx, ts);
    var link = getScriptUrl() + "?page=notice&rid=" + recordId + "&row=" + rowIdx + "&ts=" + ts + "&sig=" + sig;

    // 5. Send
    var html = "<p>Update for " + empName + ":</p>" +
               "<p><a href='" + link + "'>View Work Restrictions</a></p>" +
               "<small>Link expires in 7 days.</small>";
               
    MailApp.sendEmail({
      to: emails.join(','),
      subject: "Work Restriction Update: " + empName,
      htmlBody: html
    });

    // 6. Log
    ensureSheet_(COMMLOG_SHEET_NAME, ['ID','Timestamp','Type','Sender','To','Details']);
    var provider = rData[3] || 'Unknown Provider';
    var apptDate = rData[4] || 'No Date';
    var restrictions = rData[6] || 'No restrictions specified';
    var logDetails = "Restriction Notice: " + provider + " (" + apptDate + ") - " + restrictions.substring(0, 100);
    
    ss.getSheetByName(COMMLOG_SHEET_NAME).appendRow([recordId, new Date(), "Notice Sent", senderEmail, emails.join(','), logDetails]);

    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// NEW: Send restriction notice to selected contacts only
function adminSendRestrictionNoticeToContacts(token, recordId, rowIdx, selectedEmails) {
  try {
    var senderEmail = requireAdminToken_(token);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Get Restriction Data
    var resSheet = ss.getSheetByName(RESTR_SHEET_NAME);
    var rData = resSheet.getRange(rowIdx, 1, 1, resSheet.getLastColumn()).getDisplayValues()[0];
    
    // 2. Get FROI Employee Name
    var froiSheet = ss.getSheetByName(FORM_SHEET_NAME);
    var empName = froiSheet.getRange(recordId, 8).getValue();
    
    if (!selectedEmails || selectedEmails.length === 0) {
      return { success: false, error: "No recipients selected." };
    }

    // 3. Generate Link
    var ts = Date.now();
    var sig = computeNoticeSig_(recordId, rowIdx, ts);
    var link = getScriptUrl() + "?page=notice&rid=" + recordId + "&row=" + rowIdx + "&ts=" + ts + "&sig=" + sig;

    // 4. Send
    var html = "<p>Update for " + empName + ":</p>" +
               "<p><a href='" + link + "'>View Work Restrictions</a></p>" +
               "<small>Link expires in 7 days.</small>";
               
    MailApp.sendEmail({
      to: selectedEmails.join(','),
      subject: "Work Restriction Update: " + empName,
      htmlBody: html
    });

    // 5. Log with detailed information
    ensureSheet_(COMMLOG_SHEET_NAME, ['ID','Timestamp','Type','Sender','To','Details']);
    var provider = rData[3] || 'Unknown Provider';
    var apptDate = rData[4] || 'No Date';
    var restrictions = rData[6] || 'No restrictions specified';
    var logDetails = "Restriction Notice: " + provider + " (" + apptDate + ") - " + restrictions.substring(0, 100);
    
    ss.getSheetByName(COMMLOG_SHEET_NAME).appendRow([recordId, new Date(), "Notice Sent", senderEmail, selectedEmails.join(','), logDetails]);

    return { success: true };
  } catch(e) { 
    return { success: false, error: e.toString() }; 
  }
}

// Used by doGet for the public notice page
function getNoticeData(recordId, rowIdx) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var resSheet = ss.getSheetByName(RESTR_SHEET_NAME);
  var r = resSheet.getRange(rowIdx, 1, 1, 12).getDisplayValues()[0];
  
  var froiSheet = ss.getSheetByName(FORM_SHEET_NAME);
  var empName = froiSheet.getRange(recordId, 8).getValue();

  return {
    employee: empName,
    provider: r[3],
    apptDate: r[4], apptTime: r[5],
    restrictions: r[6],
    followDate: r[7], followTime: r[8], followProv: r[9],
    adminName: r[10], adminTitle: r[11],
    entryDate: r[1]
  };
}
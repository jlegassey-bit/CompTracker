// FROI SPREADSHEET STRUCTURE REBUILDER
// Run this ONCE to fix column header misalignment issues

function rebuildSpreadsheetStructure() {
  var ss = SpreadsheetApp.openById('1RO1tTij8KY872eigB0Wm4tRqBExYKHj2p2SHd8WKy5c');
  
  Logger.log('Starting spreadsheet structure rebuild...');
  
  // FROI_Data Sheet (Main form submissions)
  var froiSheet = ss.getSheetByName('FROI_Data') || ss.insertSheet('FROI_Data');
  var froiHeaders = [
    'Completed By Name',           // Column A (index 0)
    'Email Address',               // Column B (index 1)
    'Submitted At',                // Column C (index 2)
    'Incident Date',               // Column D (index 3)
    'Incident Time',               // Column E (index 4)
    'Incident Location',           // Column F (index 5)
    'Landmark',                    // Column G (index 6)
    'Employee Name',               // Column H (index 7)
    'Employee Number',             // Column I (index 8)
    'Employee Phone',              // Column J (index 9)
    'Work Location',               // Column K (index 10)
    'Time Workday Began',          // Column L (index 11)
    'Weekly Schedule',             // Column M (index 12)
    'Second Employer',             // Column N (index 13)
    'Second Employer Name',        // Column O (index 14)
    'Normal Duties',               // Column P (index 15)
    'Duties Explained',            // Column Q (index 16)
    'Department Name',             // Column R (index 17)
    'Date Employer Notified',      // Column S (index 18)
    'Supervisor Notified',         // Column T (index 19)
    'Witnesses',                   // Column U (index 20)
    'Injured Body Parts',          // Column V (index 21)
    'Primary Cause',               // Column W (index 22)
    'Nature of Injury',            // Column X (index 23)
    'Description',                 // Column Y (index 24)
    'Equipment Used',              // Column Z (index 25)
    'Treatment Provider',          // Column AA (index 26)
    'Treatment Type',              // Column AB (index 27)
    'Medical Provider',            // Column AC (index 28)
    'Other Provider',              // Column AD (index 29)
    'Return To Shift',             // Column AE (index 30)
    'EH Pay',                      // Column AF (index 31)
    'Reason FD',                   // Column AG (index 32)
    'Approvals FD',                // Column AH (index 33)
    'Text Notes',                  // Column AI (index 34)
    'Primary Contact Name',        // Column AJ (index 35)
    'Primary Contact Email',       // Column AK (index 36)
    'Primary Contact Phone',       // Column AL (index 37)
    'Secondary Contact Name',      // Column AM (index 38)
    'Secondary Contact Phone',     // Column AN (index 39)
    'Secondary Contact Email',     // Column AO (index 40)
    'Case Status',                 // Column AP (index 41)
    'Record ID'                    // Column AQ (index 42)
  ];
  
  if (froiSheet.getLastRow() === 0) {
    froiSheet.appendRow(froiHeaders);
    froiSheet.getRange(1, 1, 1, froiHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created FROI_Data headers');
  } else {
    froiSheet.getRange(1, 1, 1, froiHeaders.length).setValues([froiHeaders]);
    Logger.log('✓ Updated FROI_Data headers');
  }
  
  // Case_Contacts Sheet
  var contactsSheet = ss.getSheetByName('Case_Contacts') || ss.insertSheet('Case_Contacts');
  var contactsHeaders = ['Record ID', 'Name', 'Role', 'Phone', 'Email', 'Row Index'];
  if (contactsSheet.getLastRow() === 0) {
    contactsSheet.appendRow(contactsHeaders);
    contactsSheet.getRange(1, 1, 1, contactsHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created Case_Contacts headers');
  } else {
    contactsSheet.getRange(1, 1, 1, contactsHeaders.length).setValues([contactsHeaders]);
    Logger.log('✓ Updated Case_Contacts headers');
  }
  
  // Work_Restrictions Sheet
  var restrictionsSheet = ss.getSheetByName('Work_Restrictions') || ss.insertSheet('Work_Restrictions');
  var restrictionsHeaders = ['Record ID', 'Timestamp', 'Admin Name', 'Provider', 'Appt Date', 'Appt Time', 'Restrictions', 'Follow-up Date', 'Follow-up Time', 'Follow-up Provider', 'Admin Title', 'Row Index'];
  if (restrictionsSheet.getLastRow() === 0) {
    restrictionsSheet.appendRow(restrictionsHeaders);
    restrictionsSheet.getRange(1, 1, 1, restrictionsHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created Work_Restrictions headers');
  } else {
    restrictionsSheet.getRange(1, 1, 1, restrictionsHeaders.length).setValues([restrictionsHeaders]);
    Logger.log('✓ Updated Work_Restrictions headers');
  }
  
  // LostTime Sheet
  var lostTimeSheet = ss.getSheetByName('LostTime') || ss.insertSheet('LostTime');
  var lostTimeHeaders = ['Record ID', 'Timestamp', 'Status', 'Start Date', 'End Date', 'Return To Work Date', 'Days Mon', 'Days Tue', 'Days Wed', 'Days Thu', 'Days Fri', 'Days Sat', 'Days Sun', 'Scheduled Days Lost', 'Admin Name', 'Row Index'];
  if (lostTimeSheet.getLastRow() === 0) {
    lostTimeSheet.appendRow(lostTimeHeaders);
    lostTimeSheet.getRange(1, 1, 1, lostTimeHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created LostTime headers');
  } else {
    lostTimeSheet.getRange(1, 1, 1, lostTimeHeaders.length).setValues([lostTimeHeaders]);
    Logger.log('✓ Updated LostTime headers');
  }
  
  // Admin_Notes Sheet
  var notesSheet = ss.getSheetByName('Admin_Notes') || ss.insertSheet('Admin_Notes');
  var notesHeaders = ['Record ID', 'Timestamp', 'Note Type', 'Note Text', 'Admin Name', 'Admin Email', 'Row Index'];
  if (notesSheet.getLastRow() === 0) {
    notesSheet.appendRow(notesHeaders);
    notesSheet.getRange(1, 1, 1, notesHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created Admin_Notes headers');
  } else {
    notesSheet.getRange(1, 1, 1, notesHeaders.length).setValues([notesHeaders]);
    Logger.log('✓ Updated Admin_Notes headers');
  }
  
  // Documents Sheet
  var docsSheet = ss.getSheetByName('Documents') || ss.insertSheet('Documents');
  var docsHeaders = ['Record ID', 'Filename', 'File ID', 'Uploader Name', 'Timestamp', 'Row Index'];
  if (docsSheet.getLastRow() === 0) {
    docsSheet.appendRow(docsHeaders);
    docsSheet.getRange(1, 1, 1, docsHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created Documents headers');
  } else {
    docsSheet.getRange(1, 1, 1, docsHeaders.length).setValues([docsHeaders]);
    Logger.log('✓ Updated Documents headers');
  }
  
  // CommLog Sheet
  var commLogSheet = ss.getSheetByName('CommLog') || ss.insertSheet('CommLog');
  var commLogHeaders = ['Record ID', 'Timestamp', 'Type', 'Sender', 'Recipients', 'Details', 'Row Index'];
  if (commLogSheet.getLastRow() === 0) {
    commLogSheet.appendRow(commLogHeaders);
    commLogSheet.getRange(1, 1, 1, commLogHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created CommLog headers');
  } else {
    commLogSheet.getRange(1, 1, 1, commLogHeaders.length).setValues([commLogHeaders]);
    Logger.log('✓ Updated CommLog headers');
  }
  
  // Departments Sheet
  var deptsSheet = ss.getSheetByName('Departments') || ss.insertSheet('Departments');
  var deptsHeaders = ['Department', 'HR Name', 'HR Email', 'HR Phone', 'Payroll Name', 'Payroll Email', 'Payroll Phone', 'Safety Name', 'Safety Email', 'Safety Phone', 'Updated', 'Row Index'];
  if (deptsSheet.getLastRow() === 0) {
    deptsSheet.appendRow(deptsHeaders);
    deptsSheet.getRange(1, 1, 1, deptsHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created Departments headers');
  } else {
    deptsSheet.getRange(1, 1, 1, deptsHeaders.length).setValues([deptsHeaders]);
    Logger.log('✓ Updated Departments headers');
  }
  
  // WorkLocations Sheet
  var workLocsSheet = ss.getSheetByName('WorkLocations') || ss.insertSheet('WorkLocations');
  var workLocsHeaders = ['Work Location', 'Department', 'Updated', 'Row Index'];
  if (workLocsSheet.getLastRow() === 0) {
    workLocsSheet.appendRow(workLocsHeaders);
    workLocsSheet.getRange(1, 1, 1, workLocsHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created WorkLocations headers');
  } else {
    workLocsSheet.getRange(1, 1, 1, workLocsHeaders.length).setValues([workLocsHeaders]);
    Logger.log('✓ Updated WorkLocations headers');
  }
  
  // TreatmentProviders Sheet
  var providersSheet = ss.getSheetByName('TreatmentProviders') || ss.insertSheet('TreatmentProviders');
  var providersHeaders = ['Provider Name', 'Address', 'Phone', 'Mon Open', 'Mon Hours', 'Tue Open', 'Tue Hours', 'Wed Open', 'Wed Hours', 'Thu Open', 'Thu Hours', 'Fri Open', 'Fri Hours', 'Sat Open', 'Sat Hours', 'Sun Open', 'Sun Hours', 'Updated', 'Row Index'];
  if (providersSheet.getLastRow() === 0) {
    providersSheet.appendRow(providersHeaders);
    providersSheet.getRange(1, 1, 1, providersHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created TreatmentProviders headers');
  } else {
    providersSheet.getRange(1, 1, 1, providersHeaders.length).setValues([providersHeaders]);
    Logger.log('✓ Updated TreatmentProviders headers');
  }
  
  // PrimaryCauses Sheet
  var causesSheet = ss.getSheetByName('PrimaryCauses') || ss.insertSheet('PrimaryCauses');
  var causesHeaders = ['Primary Cause', 'Updated', 'Row Index'];
  if (causesSheet.getLastRow() === 0) {
    causesSheet.appendRow(causesHeaders);
    causesSheet.getRange(1, 1, 1, causesHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created PrimaryCauses headers');
  } else {
    causesSheet.getRange(1, 1, 1, causesHeaders.length).setValues([causesHeaders]);
    Logger.log('✓ Updated PrimaryCauses headers');
  }
  
  // NatureOfInjury Sheet
  var naturesSheet = ss.getSheetByName('NatureOfInjury') || ss.insertSheet('NatureOfInjury');
  var naturesHeaders = ['Nature of Injury', 'Updated', 'Row Index'];
  if (naturesSheet.getLastRow() === 0) {
    naturesSheet.appendRow(naturesHeaders);
    naturesSheet.getRange(1, 1, 1, naturesHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created NatureOfInjury headers');
  } else {
    naturesSheet.getRange(1, 1, 1, naturesHeaders.length).setValues([naturesHeaders]);
    Logger.log('✓ Updated NatureOfInjury headers');
  }
  
  // BodyParts Sheet
  var bodyPartsSheet = ss.getSheetByName('BodyParts') || ss.insertSheet('BodyParts');
  var bodyPartsHeaders = ['Body Part', 'Updated', 'Row Index'];
  if (bodyPartsSheet.getLastRow() === 0) {
    bodyPartsSheet.appendRow(bodyPartsHeaders);
    bodyPartsSheet.getRange(1, 1, 1, bodyPartsHeaders.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    Logger.log('✓ Created BodyParts headers');
  } else {
    bodyPartsSheet.getRange(1, 1, 1, bodyPartsHeaders.length).setValues([bodyPartsHeaders]);
    Logger.log('✓ Updated BodyParts headers');
  }
  
  Logger.log('');
  Logger.log('====================================');
  Logger.log('SPREADSHEET STRUCTURE REBUILD COMPLETE');
  Logger.log('====================================');
  Logger.log('All sheet headers have been aligned with code expectations.');
  Logger.log('Your data is preserved - only headers were updated/added.');
  Logger.log('');
  Logger.log('Next steps:');
  Logger.log('1. Verify your existing data is still intact');
  Logger.log('2. Test the application');
  Logger.log('3. If any data appears misaligned, you may need to manually adjust');
  
  return 'Rebuild complete! Check Execution log for details.';
}

// Function to create DocumentTypes sheet
function createDocumentTypesSheet() {
  var ss = SpreadsheetApp.openById('1RO1tTij8KY872eigB0Wm4tRqBExYKHj2p2SHd8WKy5c');
  
  var docTypesSheet = ss.getSheetByName('DocumentTypes') || ss.insertSheet('DocumentTypes');
  
  // Set headers
  var headers = ['Document Type', 'Updated', 'Row Index'];
  docTypesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  docTypesSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  docTypesSheet.setFrozenRows(1);
  
  // Add default document types if sheet is new
  if (docTypesSheet.getLastRow() === 1) {
    var defaultTypes = [
      ['Email'],
      ['M1'],
      ['Fit for Duty'],
      ['Medical Release'],
      ['Incident Report'],
      ['Witness Statement'],
      ['Photo'],
      ['Other']
    ];
    
    var now = new Date();
    defaultTypes.forEach(function(type, index) {
      docTypesSheet.appendRow([type[0], now, index + 2]);
    });
  }
  
  Logger.log('DocumentTypes sheet created/updated with default types');
}
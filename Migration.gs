// ========================================
// RECORD ID MIGRATION SCRIPT
// ========================================

function MIGRATE_TO_RECORD_IDS() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ui = SpreadsheetApp.getUi();
  
  // Safety check
  var response = ui.alert(
    'MIGRATION WARNING',
    'This will:\n\n' +
    '1. Add Record ID column to FROI Form\n' +
    '2. Generate IDs for all existing cases\n' +
    '3. Update all related sheets\n\n' +
    'MAKE SURE YOU HAVE A BACKUP!\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Migration cancelled');
    return;
  }
  
  try {
    // Step 1: Migrate FROI Form sheet
    ui.alert('Step 1/7: Migrating FROI Form sheet...');
    migrateFroiFormSheet_(ss);
    
    // Step 2: Migrate Case_Contacts
    ui.alert('Step 2/7: Migrating Case_Contacts...');
    migrateRelatedSheet_(ss, 'Case_Contacts');
    
    // Step 3: Migrate Case_Restrictions
    ui.alert('Step 3/7: Migrating Case_Restrictions...');
    migrateRelatedSheet_(ss, 'Case_Restrictions');
    
    // Step 4: Migrate Case_LostTime
    ui.alert('Step 4/7: Migrating Case_LostTime...');
    migrateRelatedSheet_(ss, 'Case_LostTime');
    
    // Step 5: Migrate Case_Docs
    ui.alert('Step 5/7: Migrating Case_Docs...');
    migrateRelatedSheet_(ss, 'Case_Docs');
    
    // Step 6: Migrate Case_CommLog and Case_Notes
    ui.alert('Step 6/7: Migrating logs and notes...');
    migrateRelatedSheet_(ss, 'Case_CommLog');
    migrateRelatedSheet_(ss, 'Case_Notes');
    
    // Step 7: Migrate Work_Schedule and Correspondence
    ui.alert('Step 7/7: Migrating work schedule and correspondence...');
    migrateRelatedSheet_(ss, 'Work_Schedule');
    migrateRelatedSheet_(ss, 'Case_Correspondence');
    
    ui.alert('✅ MIGRATION COMPLETE!\n\nAll sheets have been updated with Record IDs.');
    
  } catch (error) {
    ui.alert('❌ ERROR: ' + error.toString() + '\n\nRestore from backup!');
    Logger.log('Migration error: ' + error.toString());
  }
}

function migrateFroiFormSheet_(ss) {
  var sheet = ss.getSheetByName(FORM_SHEET_NAME);
  if (!sheet) throw new Error('FROI Form sheet not found');
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data to migrate
  
  // Insert new column A for Record ID
  sheet.insertColumnBefore(1);
  
  // Set header
  sheet.getRange(1, 1).setValue('Record ID');
  
  // Generate Record IDs for all existing rows
  var recordIds = [];
  var existingIds = new Set();
  
  for (var i = 2; i <= lastRow; i++) {
    var newId = generateUniqueRecordId_(existingIds);
    recordIds.push([newId]);
    existingIds.add(newId);
    
    // Store mapping for related sheets
    PropertiesService.getScriptProperties().setProperty('ROW_' + i, newId);
  }
  
  // Write all Record IDs at once
  sheet.getRange(2, 1, recordIds.length, 1).setValues(recordIds);
  
  Logger.log('Migrated ' + recordIds.length + ' FROI records');
}

function migrateRelatedSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('Sheet ' + sheetName + ' not found or empty, skipping');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();
  
  // Update column A (ID column) with Record IDs
  for (var i = 1; i < data.length; i++) {
    var oldRowNum = data[i][0];
    
    // Try to get mapped Record ID
    var recordId = props.getProperty('ROW_' + oldRowNum);
    
    if (recordId) {
      data[i][0] = recordId;
    } else {
      Logger.log('Warning: No mapping found for row ' + oldRowNum + ' in ' + sheetName);
    }
  }
  
  // Write updated data back
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  Logger.log('Migrated ' + sheetName);
}

function generateUniqueRecordId_(existingIds) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var maxAttempts = 100;
  
  for (var attempt = 0; attempt < maxAttempts; attempt++) {
    var id = 'FROI-';
    for (var i = 0; i < 10; i++) {
      id += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    
    if (!existingIds.has(id)) {
      return id;
    }
  }
  
  throw new Error('Could not generate unique Record ID after ' + maxAttempts + ' attempts');
}

// Cleanup function - run this after migration is complete
function CLEANUP_MIGRATION_PROPERTIES() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  
  keys.forEach(function(key) {
    if (key.startsWith('ROW_')) {
      props.deleteProperty(key);
    }
  });
  
  SpreadsheetApp.getUi().alert('Migration properties cleaned up');
}
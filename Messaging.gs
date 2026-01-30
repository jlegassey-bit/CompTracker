// ========================================
// MESSAGING FUNCTIONS FOR FROI SYSTEM
// ========================================

/**
 * Get messaging contacts for a case (employee, supervisor, case contacts)
 */
function adminGetMessagingContacts(token, recordId) {
  try {
    requireAdminToken_(token);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var formSheet = ss.getSheetByName(FORM_SHEET_NAME);
    var contactsSheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
    
    if (!formSheet || !contactsSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    var recipients = [];
    
    // Find FROI data by searching for Record ID in column A
    var froiData = formSheet.getDataRange().getValues();
    var froiRow = null;
    
    for (var i = 1; i < froiData.length; i++) {
      if (String(froiData[i][0]) === String(recordId)) {
        froiRow = froiData[i];
        break;
      }
    }
    
    if (!froiRow) {
      return { success: false, error: 'Case not found' };
    }
    
    // Add employee if email exists (Column B - index 1)
    var empEmail = safe_(froiRow[1]);
    var empName = safe_(froiRow[7]); // Column H
    if (empEmail && empEmail.indexOf('@') > -1) {
      recipients.push({
        id: 'employee_' + recordId,
        name: empName || 'Employee',
        email: empEmail,
        role: 'Injured Employee',
        type: 'employee'
      });
    }
    
    // Add supervisor if exists
    var supName = safe_(froiRow[11]); // Column L
    var supEmail = safe_(froiRow[19]); // Column T
    if (supEmail && supEmail.indexOf('@') > -1) {
      recipients.push({
        id: 'supervisor_' + recordId,
        name: supName || 'Supervisor',
        email: supEmail,
        role: 'Supervisor',
        type: 'supervisor'
      });
    }
    
    // Get case contacts
    var contactsData = contactsSheet.getDataRange().getValues();
    for (var i = 1; i < contactsData.length; i++) {
      if (String(contactsData[i][0]) === String(recordId)) {
        var email = safe_(contactsData[i][4]);
        if (email && email.indexOf('@') > -1) {
          recipients.push({
            id: 'contact_' + i,
            name: safe_(contactsData[i][1]),
            email: email,
            role: safe_(contactsData[i][2]) || 'Case Contact',
            type: 'contact'
          });
        }
      }
    }
    
    return { success: true, recipients: recipients };
    
  } catch (error) {
    Logger.log('Error in adminGetMessagingContacts: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Send case email and log to correspondence
 */
function adminSendCaseEmail(token, recordId, recipientEmails, subject, message, attachmentData) {
  try {
    var senderEmail = requireAdminToken_(token);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    var corrSheet = ensureSheet_('Case_Correspondence', [
      'Record ID', 'Timestamp', 'Sender', 'Recipients', 'Subject', 
      'Message', 'Attachments', 'Type', 'Email ID', 'Attachment IDs'
    ]);
    
    // Get sender name
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var senderName = 'Admin User';
    if (adminSheet && adminSheet.getLastRow() > 1) {
      var admins = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < admins.length; i++) {
        if (normalizeEmail_(admins[i][0]) === normalizeEmail_(senderEmail)) {
          senderName = admins[i][1] || 'Admin User';
          break;
        }
      }
    }
    
    // Add case ID tag
    var subjectWithTag = subject;
    if (subject.indexOf('[' + recordId + ']') === -1) {
      subjectWithTag = '[' + recordId + '] ' + subject;
    }
    
    // Format email
    var htmlBody = formatEmailBody_(message, recordId);
    
    var emailOptions = {
      name: 'FROI System - ' + senderName,
      replyTo: senderEmail,
      htmlBody: htmlBody
    };
    
    // Process attachments
    var attachmentNames = [];
    var attachmentIds = [];
    
    if (attachmentData && attachmentData.length > 0) {
      var blobs = [];
      var docsSheet = ensureSheet_(DOCS_SHEET_NAME, ['ID','Filename','FileID','UploadedBy','Timestamp']);
      
      var folder;
      var folders = DriveApp.getFoldersByName("FROI_Uploads");
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder("FROI_Uploads");
      }
      
      for (var i = 0; i < attachmentData.length; i++) {
        var att = attachmentData[i];
        var blob = Utilities.newBlob(
          Utilities.base64Decode(att.data),
          att.mimeType,
          att.name
        );
        blobs.push(blob);
        attachmentNames.push(att.name);
        
        try {
          var file = folder.createFile(blob);
          var fileId = file.getId();
          attachmentIds.push(fileId);
          
          docsSheet.appendRow([
            recordId,
            att.name,
            fileId,
            senderName + ' (Email)',
            new Date()
          ]);
        } catch (uploadErr) {
          Logger.log('Failed to upload attachment: ' + uploadErr);
        }
      }
      
      emailOptions.attachments = blobs;
    }
    
    // Send email
    var recipientList = recipientEmails.join(', ');
    MailApp.sendEmail(recipientList, subjectWithTag, message, emailOptions);
    
    // Log to correspondence
    var timestamp = new Date();
    var emailId = 'SENT_' + timestamp.getTime();
    
    corrSheet.appendRow([
      recordId,
      timestamp,
      senderEmail,
      recipientList,
      subjectWithTag,
      message,
      attachmentNames.join(', '),
      'Sent',
      emailId,
      attachmentIds.join(', ')
    ]);
    
    // Log to CommLog
    var commSheet = ensureSheet_('Case_CommLog', ['ID','Timestamp','Type','Sender','To','Details']);
    commSheet.appendRow([
      recordId,
      timestamp,
      'Email Sent',
      senderEmail,
      recipientList,
      'Subject: ' + subjectWithTag
    ]);
    
    return { 
      success: true, 
      message: 'Email sent successfully',
      emailId: emailId
    };
    
  } catch (error) {
    Logger.log('Error in adminSendCaseEmail: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Format email body
 */
function formatEmailBody_(message, recordId) {
  var html = '<div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">';
  html += '<div style="background: #343a40; color: white; padding: 20px; border-radius: 8px 8px 0 0;">';
  html += '<h2 style="margin: 0; font-size: 20px;">City of Portland - FROI Management</h2>';
  html += '<p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Case: ' + recordId + '</p>';
  html += '</div>';
  html += '<div style="background: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; border-top: none;">';
  html += '<div style="background: white; padding: 20px; border-radius: 6px; border: 1px solid #e0e0e0;">';
  html += message.replace(/\n/g, '<br>');
  html += '</div>';
  html += '<div style="margin-top: 20px; padding-top: 15px; border-top: 2px solid #007bff; font-size: 11px; color: #6c757d;">';
  html += '<p style="margin: 0;"><strong>Sent from FROI Management System</strong></p>';
  html += '<p style="margin: 5px 0 0 0;">Please include case ID <strong>[' + recordId + ']</strong> in any replies.</p>';
  html += '<p style="margin: 10px 0 0 0; font-style: italic;">City of Portland, Maine - Human Resources Department</p>';
  html += '</div>';
  html += '</div>';
  html += '</div>';
  return html;
}

/**
 * Attach reply to case
 */
function adminAttachReply(token, recordId, emailId) {
  try {
    requireAdminToken_(token);
    
    var message = GmailApp.getMessageById(emailId);
    if (!message) {
      return { success: false, error: 'Email message not found' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var corrSheet = ensureSheet_('Case_Correspondence', [
      'Record ID', 'Timestamp', 'Sender', 'Recipients', 'Subject', 
      'Message', 'Attachments', 'Type', 'Email ID', 'Attachment IDs'
    ]);
    
    // Check if already attached
    var corrData = corrSheet.getDataRange().getValues();
    for (var i = 1; i < corrData.length; i++) {
      if (corrData[i][8] === emailId) {
        return { success: false, error: 'This reply has already been attached' };
      }
    }
    
    // Process attachments
    var attachments = message.getAttachments();
    var attachmentNames = [];
    var attachmentIds = [];
    var senderName = message.getFrom();
    
    if (attachments.length > 0) {
      var docsSheet = ensureSheet_(DOCS_SHEET_NAME, ['ID','Filename','FileID','UploadedBy','Timestamp']);
      
      var folder;
      var folders = DriveApp.getFoldersByName("FROI_Uploads");
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder("FROI_Uploads");
      }
      
      for (var a = 0; a < attachments.length; a++) {
        var att = attachments[a];
        attachmentNames.push(att.getName());
        
        try {
          var file = folder.createFile(att);
          var fileId = file.getId();
          attachmentIds.push(fileId);
          
          docsSheet.appendRow([
            recordId,
            att.getName(),
            fileId,
            senderName + ' (Reply)',
            new Date()
          ]);
        } catch (uploadErr) {
          Logger.log('Failed to upload reply attachment: ' + uploadErr);
        }
      }
    }
    
    // Add to correspondence
    corrSheet.appendRow([
      recordId,
      message.getDate(),
      message.getFrom(),
      message.getTo(),
      message.getSubject(),
      message.getPlainBody(),
      attachmentNames.join(', '),
      'Reply',
      emailId,
      attachmentIds.join(', ')
    ]);
    
    // Log to CommLog
    var commSheet = ensureSheet_('Case_CommLog', ['ID','Timestamp','Type','Sender','To','Details']);
    commSheet.appendRow([
      recordId,
      new Date(),
      'Email Reply Attached',
      Session.getActiveUser().getEmail(),
      '',
      'Attached reply from: ' + message.getFrom()
    ]);
    
    return { success: true, message: 'Reply attached successfully' };
    
  } catch (error) {
    Logger.log('Error in adminAttachReply: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete correspondence
 */
function adminDeleteCorrespondence(token, recordId, emailId) {
  try {
    requireAdminToken_(token);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var corrSheet = ss.getSheetByName('Case_Correspondence');
    
    if (!corrSheet) {
      return { success: false, error: 'Correspondence sheet not found' };
    }
    
    var data = corrSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(recordId) && data[i][8] === emailId) {
        corrSheet.deleteRow(i + 1);
        
        var commSheet = ensureSheet_('Case_CommLog', ['ID','Timestamp','Type','Sender','To','Details']);
        commSheet.appendRow([
          recordId,
          new Date(),
          'Correspondence Deleted',
          Session.getActiveUser().getEmail(),
          '',
          'Deleted email record: ' + emailId
        ]);
        
        return { success: true, message: 'Correspondence deleted' };
      }
    }
    
    return { success: false, error: 'Correspondence not found' };
    
  } catch (error) {
    Logger.log('Error in adminDeleteCorrespondence: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get correspondence for a case
 */
function getCaseCorrespondence_(recordId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var corrSheet = ss.getSheetByName('Case_Correspondence');
    
    if (!corrSheet || corrSheet.getLastRow() < 2) {
      return [];
    }
    
    var rid = String(recordId);
    var data = corrSheet.getRange(2, 1, corrSheet.getLastRow() - 1, 10).getValues();
    var correspondence = [];
    
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === rid) {
        var timestamp = data[i][1];
        if (timestamp instanceof Date) {
          timestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        } else {
          timestamp = String(timestamp);
        }
        
        correspondence.push({
          recordId: String(data[i][0]),
          timestamp: timestamp,
          sender: String(data[i][2]),
          recipients: String(data[i][3]),
          subject: String(data[i][4]),
          message: String(data[i][5]),
          attachments: String(data[i][6]),
          type: String(data[i][7]),
          emailId: data[i][8],
          attachmentIds: String(data[i][9] || '')
        });
      }
    }
    
    correspondence.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    
    return correspondence;
    
  } catch (error) {
    Logger.log('Error in getCaseCorrespondence_: ' + error.toString());
    return [];
  }
}
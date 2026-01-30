// --- GENERIC CRUD HELPERS ---
function genericList_(token, sheetName, minCols) {
  requireAdminToken_(token);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), minCols)).getDisplayValues();
  return data.map(function(row, idx) { return { row: row, rowIdx: idx + 2 }; });
}

function genericAdd_(token, sheetName, rowArray) {
  requireAdminToken_(token);
  var sh = ensureSheet_(sheetName, []);
  rowArray.push(new Date()); // Add Updated timestamp
  sh.appendRow(rowArray);
  return { success: true };
}

function genericUpdate_(token, sheetName, rowIdx, rowArray) {
  requireAdminToken_(token);
  var r = parseInt(rowIdx, 10);
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sh || r < 2 || r > sh.getLastRow()) return { success: false, error: 'Invalid row' };
  
  rowArray.push(new Date()); // Update timestamp
  sh.getRange(r, 1, 1, rowArray.length).setValues([rowArray]);
  return { success: true };
}

function genericDelete_(token, sheetName, rowIdx) {
  requireAdminToken_(token);
  var r = parseInt(rowIdx, 10);
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sh) return { success: false, error: 'Sheet not found' };
  sh.deleteRow(r);
  return { success: true };
}
// --- SEND WELCOME EMAIL TO NEW ADMIN ---
function sendAdminWelcomeEmail_(adminData, addedByEmail) {
  try {
    var appUrl = ScriptApp.getService().getUrl();
    var currentYear = new Date().getFullYear();
    
    var htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f6f9; padding: 40px 20px;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
          
          <!-- Header with City Branding -->
          <tr>
            <td style="background: linear-gradient(90deg, #0b2a4a 0%, #0f3b67 50%, #164b83 100%); padding: 30px; text-align: center; border-radius: 8px 8px 0 0;">
              <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: 700; letter-spacing: 0.5px;">
                City of Portland
              </h1>
              <p style="margin: 8px 0 0 0; color: #e0e7ef; font-size: 14px; font-weight: 400;">
                Human Resources Department
              </p>
            </td>
          </tr>
          
          <!-- Welcome Message -->
          <tr>
            <td style="padding: 40px 40px 20px 40px;">
              <h2 style="margin: 0 0 20px 0; color: #0f3b67; font-size: 22px; font-weight: 600;">
                Welcome to the FROI Management System
              </h2>
              <p style="margin: 0 0 16px 0; color: #333333; font-size: 16px; line-height: 1.6;">
                Hello <strong>${adminData.name}</strong>,
              </p>
              <p style="margin: 0 0 16px 0; color: #555555; font-size: 15px; line-height: 1.6;">
                You have been granted administrator access to the City of Portland's First Report of Injury (FROI) Management System. This system allows you to manage workplace injury reports, track case progress, and coordinate with departments across the organization.
              </p>
            </td>
          </tr>
          
          <!-- Account Details -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f8f9fa; border-left: 4px solid #0f3b67; border-radius: 4px;">
                <tr>
                  <td style="padding: 20px;">
                    <p style="margin: 0 0 12px 0; color: #666666; font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">
                      Your Account Details
                    </p>
                    <table width="100%" cellpadding="4" cellspacing="0">
                      <tr>
                        <td style="color: #555555; font-size: 14px; padding: 4px 0;"><strong>Name:</strong></td>
                        <td style="color: #333333; font-size: 14px; padding: 4px 0;">${adminData.name}</td>
                      </tr>
                      <tr>
                        <td style="color: #555555; font-size: 14px; padding: 4px 0;"><strong>Email:</strong></td>
                        <td style="color: #333333; font-size: 14px; padding: 4px 0;">${adminData.email}</td>
                      </tr>
                      <tr>
                        <td style="color: #555555; font-size: 14px; padding: 4px 0;"><strong>Department:</strong></td>
                        <td style="color: #333333; font-size: 14px; padding: 4px 0;">${adminData.department}</td>
                      </tr>
                      ${adminData.title ? `
                      <tr>
                        <td style="color: #555555; font-size: 14px; padding: 4px 0;"><strong>Title:</strong></td>
                        <td style="color: #333333; font-size: 14px; padding: 4px 0;">${adminData.title}</td>
                      </tr>
                      ` : ''}
                      <tr>
                        <td style="color: #555555; font-size: 14px; padding: 4px 0;"><strong>Access Level:</strong></td>
                        <td style="color: #333333; font-size: 14px; padding: 4px 0;">Full Administrator</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Access Button -->
          <tr>
            <td style="padding: 0 40px 30px 40px; text-align: center;">
              <a href="${appUrl}" style="display: inline-block; padding: 14px 40px; background: linear-gradient(135deg, #0f3b67 0%, #164b83 100%); color: #ffffff; text-decoration: none; border-radius: 6px; font-size: 16px; font-weight: 600; box-shadow: 0 4px 12px rgba(15, 59, 103, 0.3);">
                Access FROI System
              </a>
            </td>
          </tr>
          
          <!-- Quick Start Guide -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background: linear-gradient(135deg, #e3f2fd 0%, #f0f4f8 100%); border-left: 4px solid #2196F3; border-radius: 4px;">
                <tr>
                  <td style="padding: 24px;">
                    <h3 style="margin: 0 0 16px 0; color: #0f3b67; font-size: 18px; font-weight: 600;">
                      📘 Quick Start Guide
                    </h3>
                    
                    <!-- Step 1 -->
                    <div style="margin-bottom: 20px;">
                      <p style="margin: 0 0 8px 0; color: #0f3b67; font-size: 15px; font-weight: 600;">
                        1. Accessing the System
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • Click the "Access FROI System" button above or visit the system URL
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • You'll be authenticated automatically using your City of Portland email
                      </p>
                      <p style="margin: 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • The main dashboard will display all injury reports in a sortable table
                      </p>
                    </div>
                    
                    <!-- Step 2 -->
                    <div style="margin-bottom: 20px;">
                      <p style="margin: 0 0 8px 0; color: #0f3b67; font-size: 15px; font-weight: 600;">
                        2. Viewing & Managing Cases
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • <strong>Dashboard:</strong> Search, filter, and sort all injury reports by status, date, or employee
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • <strong>Case Details:</strong> Click any row to open the complete case file with 8 organized tabs
                      </p>
                      <p style="margin: 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • <strong>Status Updates:</strong> Use the dropdown in the case header to mark cases as Pending, Reviewed, or Closed
                      </p>
                    </div>
                    
                    <!-- Step 3 -->
                    <div style="margin-bottom: 20px;">
                      <p style="margin: 0 0 8px 0; color: #0f3b67; font-size: 15px; font-weight: 600;">
                        3. Key Tabs & Features
                      </p>
                      <table width="100%" cellpadding="6" cellspacing="0" style="background-color: #ffffff; border-radius: 4px; margin-top: 8px;">
                        <tr>
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">FROI Record</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">View and edit incident details</td>
                        </tr>
                        <tr style="background-color: #f8f9fa;">
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Work Info</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Configure employee schedule for lost-time calculations</td>
                        </tr>
                        <tr>
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Contacts</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Add supervisors, HR, safety officers, payroll contacts</td>
                        </tr>
                        <tr style="background-color: #f8f9fa;">
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Restrictions</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Log work restrictions and medical appointments</td>
                        </tr>
                        <tr>
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Lost Time</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Track time off with automatic calculation of work days lost</td>
                        </tr>
                        <tr style="background-color: #f8f9fa;">
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Documents</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Upload medical records, reports, photos securely</td>
                        </tr>
                        <tr>
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Admin Notes</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Add safety follow-up notes, investigation details</td>
                        </tr>
                        <tr style="background-color: #f8f9fa;">
                          <td style="color: #0f3b67; font-size: 13px; font-weight: 600; padding: 6px 12px;">Comm Log</td>
                          <td style="color: #555555; font-size: 13px; padding: 6px 12px;">Audit trail of all system actions and emails</td>
                        </tr>
                      </table>
                    </div>
                    
                    <!-- Step 4 -->
                    <div style="margin-bottom: 20px;">
                      <p style="margin: 0 0 8px 0; color: #0f3b67; font-size: 15px; font-weight: 600;">
                        4. Common Workflows
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        <strong>New Injury Report:</strong> Review FROI → Add Contacts → Configure Work Schedule → Track Restrictions & Lost Time
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        <strong>Medical Appointment:</strong> Open Case → Restrictions Tab → Add Provider/Date → Click "Send Notice" to email contacts
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        <strong>Safety Investigation:</strong> Open Case → Admin Notes Tab → Add "Safety Follow Up" note with findings
                      </p>
                      <p style="margin: 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        <strong>Generate Audit Report:</strong> Open Case → Click "Full Audit PDF" button in header to print comprehensive record
                      </p>
                    </div>
                    
                    <!-- Step 5 -->
                    <div style="margin-bottom: 20px;">
                      <p style="margin: 0 0 8px 0; color: #0f3b67; font-size: 15px; font-weight: 600;">
                        5. Settings & Configuration
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • Access via the ⚙️ Settings button in the top navigation
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • Manage system administrators and their access levels
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • Configure departments and their contact information
                      </p>
                      <p style="margin: 0 0 4px 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • Add work locations, injury causes, body parts, and treatment providers
                      </p>
                      <p style="margin: 0; color: #444444; font-size: 14px; line-height: 1.5;">
                        • View public-facing treatment provider directory
                      </p>
                    </div>
                    
                    <!-- Best Practices -->
                    <div style="background-color: #fff3cd; padding: 12px; border-radius: 4px; border-left: 3px solid #ffc107;">
                      <p style="margin: 0 0 8px 0; color: #856404; font-size: 14px; font-weight: 600;">
                        💡 Best Practices
                      </p>
                      <p style="margin: 0 0 4px 0; color: #856404; font-size: 13px; line-height: 1.5;">
                        • Always configure the employee's work schedule before tracking lost time
                      </p>
                      <p style="margin: 0 0 4px 0; color: #856404; font-size: 13px; line-height: 1.5;">
                        • Add all relevant contacts immediately to ensure proper communication
                      </p>
                      <p style="margin: 0 0 4px 0; color: #856404; font-size: 13px; line-height: 1.5;">
                        • Use Admin Notes to document safety investigations and administrative actions
                      </p>
                      <p style="margin: 0; color: #856404; font-size: 13px; line-height: 1.5;">
                        • Generate audit PDFs regularly for record-keeping and compliance
                      </p>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Capabilities Section -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <p style="margin: 0 0 16px 0; color: #555555; font-size: 15px; line-height: 1.6;">
                As an administrator, you have full access to:
              </p>
              <ul style="margin: 0; padding-left: 20px; color: #555555; font-size: 14px; line-height: 1.8;">
                <li>View and manage all injury reports across departments</li>
                <li>Track work restrictions and lost time with automatic calculations</li>
                <li>Manage case documentation and communication logs</li>
                <li>Add safety and administrative follow-up notes</li>
                <li>Generate audit reports and export data</li>
                <li>Configure system settings and user access</li>
              </ul>
            </td>
          </tr>
          
          <!-- Support Information -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #d1ecf1; border-left: 4px solid #17a2b8; border-radius: 4px;">
                <tr>
                  <td style="padding: 16px;">
                    <p style="margin: 0 0 8px 0; color: #0c5460; font-size: 14px; font-weight: 600;">
                      📋 Need Help?
                    </p>
                    <p style="margin: 0; color: #0c5460; font-size: 13px; line-height: 1.5;">
                      If you have questions or need assistance accessing the system, please contact the HR Technology team or reply to this email. We're here to help you get started!
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Footer -->
          <tr>
            <td style="padding: 30px 40px; background-color: #f8f9fa; border-radius: 0 0 8px 8px; border-top: 1px solid #e9ecef;">
              <p style="margin: 0 0 8px 0; color: #666666; font-size: 12px; line-height: 1.5; text-align: center;">
                This is an automated message from the City of Portland FROI Management System.
              </p>
              <p style="margin: 0; color: #999999; font-size: 11px; text-align: center;">
                © ${currentYear} City of Portland, Maine • Human Resources Department
              </p>
            </td>
          </tr>
          
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `;
    
    var plainTextBody = `
City of Portland - FROI Management System
Welcome New Administrator

Hello ${adminData.name},

You have been granted administrator access to the City of Portland's First Report of Injury (FROI) Management System.

YOUR ACCOUNT DETAILS:
- Name: ${adminData.name}
- Email: ${adminData.email}
- Department: ${adminData.department}
${adminData.title ? '- Title: ' + adminData.title : ''}
- Access Level: Full Administrator

Access the system here: ${appUrl}

QUICK START GUIDE:

1. Accessing the System
   • Click the link above or visit the system URL
   • You'll be authenticated automatically using your City of Portland email
   • The main dashboard will display all injury reports

2. Viewing & Managing Cases
   • Dashboard: Search, filter, and sort all injury reports
   • Case Details: Click any row to open the complete case file
   • Status Updates: Mark cases as Pending, Reviewed, or Closed

3. Key Tabs & Features
   • FROI Record: View and edit incident details
   • Work Info: Configure employee schedule for lost-time calculations
   • Contacts: Add supervisors, HR, safety officers, payroll contacts
   • Restrictions: Log work restrictions and medical appointments
   • Lost Time: Track time off with automatic calculations
   • Documents: Upload medical records securely
   • Admin Notes: Add safety follow-up notes
   • Comm Log: Audit trail of all system actions

4. Common Workflows
   • New Injury Report: Review FROI → Add Contacts → Configure Schedule → Track Time
   • Medical Appointment: Restrictions Tab → Add Provider → Send Notice
   • Safety Investigation: Admin Notes Tab → Add "Safety Follow Up" note
   • Generate Audit: Click "Full Audit PDF" button in case header

5. Settings & Configuration
   • Access via Settings button in top navigation
   • Manage administrators, departments, work locations
   • Configure injury causes, body parts, treatment providers

BEST PRACTICES:
- Configure work schedule before tracking lost time
- Add all contacts immediately for proper communication
- Use Admin Notes to document investigations
- Generate audit PDFs regularly for compliance

If you have questions or need assistance, please contact the HR Technology team.

---
City of Portland, Maine
Human Resources Department
    `;
    
    MailApp.sendEmail({
      to: adminData.email,
      subject: 'Welcome to City of Portland FROI Management System - Administrator Access Granted',
      body: plainTextBody,
      htmlBody: htmlBody,
      name: 'City of Portland HR'
    });
    
    return { success: true };
  } catch(e) {
    Logger.log('Error sending welcome email: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// --- ADMINS ---
function settingsGetAdmins(token) {
  try {
    var list = genericList_(token, ADMIN_SHEET_NAME, 4);
    var rows = list.map(function(item) {
      return { rowIdx: item.rowIdx, email: item.row[0], name: item.row[1], title: item.row[2], department: item.row[3] };
    });
    return { success: true, rows: rows };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function settingsAddAdmin(token, obj) {
  try {
    requireAdminToken_(token);
    
    // Validate email domain
    if (!isAllowedDomain_(obj.email)) {
      throw new Error('Invalid email domain. Must be a City of Portland email.');
    }
    
    // Check for duplicate email
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(ADMIN_SHEET_NAME);
    
    if (sh && sh.getLastRow() > 1) {
      var existingEmails = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < existingEmails.length; i++) {
        if (String(existingEmails[i][0]).toLowerCase().trim() === obj.email.toLowerCase().trim()) {
          return { success: false, error: 'This email already has admin access' };
        }
      }
    }
    
    // Add new admin: [email, name, title, department]
    return genericAdd_(token, ADMIN_SHEET_NAME, [obj.email, obj.name, obj.title || '', obj.department]); 
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  }
}

function settingsUpdateAdmin(token, rowIdx, obj) {
  try { 
    return genericUpdate_(token, ADMIN_SHEET_NAME, rowIdx, [obj.email, obj.name, obj.title, obj.department]); 
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  }
}

function settingsDeleteAdmin(token, rowIdx) {
  try {
    var list = genericList_(token, ADMIN_SHEET_NAME, 1);
    if (list.length <= 1) return { success: false, error: 'Cannot delete last admin.' };
    return genericDelete_(token, ADMIN_SHEET_NAME, rowIdx);
  } catch (e) { return { success: false, error: e.toString() }; }
}

// --- DEPARTMENTS ---
function settingsGetDepartments(token) {
  try {
    var list = genericList_(token, DEPTS_SHEET_NAME, 11);
    var rows = list.map(function(item) {
      return {
        row: item.rowIdx, department: item.row[0],
        hrName: item.row[1], hrEmail: item.row[2], hrPhone: item.row[3],
        payrollName: item.row[4], payrollEmail: item.row[5], payrollPhone: item.row[6],
        safetyName: item.row[7], safetyEmail: item.row[8], safetyPhone: item.row[9], updated: item.row[10]
      };
    });
    return { success: true, rows: rows };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function settingsAddDepartment(token, name) {
  try { return genericAdd_(token, DEPTS_SHEET_NAME, [name, '','','','','','','','','']); } 
  catch (e) { return { success: false, error: e.toString() }; }
}

function settingsUpdateDepartment(token, rowIdx, name) {
  try {
    requireAdminToken_(token);
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DEPTS_SHEET_NAME);
    sh.getRange(rowIdx, 1).setValue(name);
    sh.getRange(rowIdx, 11).setValue(new Date());
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function settingsDeleteDepartment(token, rowIdx) {
  try { return genericDelete_(token, DEPTS_SHEET_NAME, rowIdx); } 
  catch (e) { return { success: false, error: e.toString() }; }
}

function settingsGetDeptContacts(token, deptName) {
  try {
    var all = settingsGetDepartments(token).rows;
    var d = all.find(function(x) { return x.department === deptName; });
    if (!d) return { success: false, error: 'Department not found' };
    return { success: true, contacts: {
        "PAO/HR Generalist": { name: d.hrName, email: d.hrEmail, phone: d.hrPhone },
        "Payroll Specialist/Manager": { name: d.payrollName, email: d.payrollEmail, phone: d.payrollPhone },
        "Safety and Training Officer": { name: d.safetyName, email: d.safetyEmail, phone: d.safetyPhone }
    }};
  } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsSaveDeptContacts(token, deptName, payload) {
  try {
    requireAdminToken_(token);
    var all = settingsGetDepartments(token).rows;
    var d = all.find(function(x) { return x.department === deptName; });
    if (!d) return { success: false, error: 'Department not found' };
    var hr = payload["PAO/HR Generalist"] || {};
    var pr = payload["Payroll Specialist/Manager"] || {};
    var sa = payload["Safety and Training Officer"] || {};
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DEPTS_SHEET_NAME);
    sh.getRange(d.row, 2, 1, 10).setValues([[hr.name, hr.email, hr.phone, pr.name, pr.email, pr.phone, sa.name, sa.email, sa.phone, new Date()]]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// --- WORK LOCATIONS ---
function settingsGetWorkLocs(token) {
  try {
    var list = genericList_(token, WORKLOCS_SHEET_NAME, 8);
    var rows = list.map(function(item) { 
      return { 
        rowIdx: item.rowIdx, 
        code: item.row[0],
        workLocation: item.row[1], 
        address: item.row[2],
        city: item.row[3],
        state: item.row[4],
        zip: item.row[5],
        department: item.row[6], 
        updated: item.row[7] 
      }; 
    });
    return { success: true, rows: rows };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function settingsAddWorkLoc(token, obj) {
  try { return genericAdd_(token, WORKLOCS_SHEET_NAME, [obj.code, obj.name, obj.address, obj.city, obj.state, obj.zip, obj.dept]); } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsUpdateWorkLoc(token, row, obj) {
  try { return genericUpdate_(token, WORKLOCS_SHEET_NAME, row, [obj.code, obj.name, obj.address, obj.city, obj.state, obj.zip, obj.dept]); } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsDeleteWorkLoc(token, row) { return genericDelete_(token, WORKLOCS_SHEET_NAME, row); }

// --- PRIMARY CAUSES ---
function settingsGetPrimaryCauses(token) {
  try {
    var list = genericList_(token, PRIMARYCAUSE_SHEET_NAME, 2);
    var rows = list.map(function(item) { return { rowIdx: item.rowIdx, cause: item.row[0], updated: item.row[1] }; });
    return { success: true, rows: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsAddPrimaryCause(token, val) { return genericAdd_(token, PRIMARYCAUSE_SHEET_NAME, [val]); }

function settingsDeletePrimaryCause(token, row) { return genericDelete_(token, PRIMARYCAUSE_SHEET_NAME, row); }

// --- TREATMENTS ---
function settingsGetTreatments(token) {
  try {
    var list = genericList_(token, TREATMENTS_SHEET_NAME, 2);
    var rows = list.map(function(item) { return { rowIdx: item.rowIdx, treatment: item.row[0], updated: item.row[1] }; });
    return { success: true, rows: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsAddTreatment(token, val) { return genericAdd_(token, TREATMENTS_SHEET_NAME, [val]); }

function settingsDeleteTreatment(token, row) { return genericDelete_(token, TREATMENTS_SHEET_NAME, row); }

// --- BODY PARTS ---
function settingsGetBodyParts(token) {
  try {
    var list = genericList_(token, BODYPARTS_SHEET_NAME, 2);
    var rows = list.map(function(item) { return { rowIdx: item.rowIdx, bodyPart: item.row[0], updated: item.row[1] }; });
    return { success: true, rows: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsAddBodyPart(token, val) { return genericAdd_(token, BODYPARTS_SHEET_NAME, [val]); }

function settingsDeleteBodyPart(token, row) { return genericDelete_(token, BODYPARTS_SHEET_NAME, row); }


// --- INJURY TYPES ---
function settingsGetInjuryTypes(token) {
  try {
    var list = genericList_(token, 'Type', 2);
    var rows = list.map(function(item) { return { rowIdx: item.rowIdx, injuryType: item.row[0], updated: item.row[1] }; });
    return { success: true, rows: rows };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function settingsAddInjuryType(token, val) {
  try { return genericAdd_(token, 'Type', [val]); } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsDeleteInjuryType(token, row) { return genericDelete_(token, 'Type', row); }

// --- PROVIDERS ---
function settingsGetProviders(token) {
  try {
    var list = genericList_(token, PROVIDERS_SHEET_NAME, 18);
    var rows = list.map(function(item) {
      var r = item.row;
      return {
        rowIdx: item.rowIdx, providerName: r[0], address: r[1], phone: r[2],
        monOpen: parseBool_(r[3]), monHours: r[4], tueOpen: parseBool_(r[5]), tueHours: r[6],
        wedOpen: parseBool_(r[7]), wedHours: r[8], thuOpen: parseBool_(r[9]), thuHours: r[10],
        friOpen: parseBool_(r[11]), friHours: r[12], satOpen: parseBool_(r[13]), satHours: r[14],
        sunOpen: parseBool_(r[15]), sunHours: r[16], updated: r[17]
      };
    });
    return { success: true, rows: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsAddProvider(token, obj) {
  var row = [obj.providerName, obj.address, obj.phone, obj.monOpen, obj.monHours, obj.tueOpen, obj.tueHours, obj.wedOpen, obj.wedHours, obj.thuOpen, obj.thuHours, obj.friOpen, obj.friHours, obj.satOpen, obj.satHours, obj.sunOpen, obj.sunHours];
  return genericAdd_(token, PROVIDERS_SHEET_NAME, row);
}

function settingsDeleteProvider(token, row) { return genericDelete_(token, PROVIDERS_SHEET_NAME, row); }
function settingsUpdateProvider(token, row, obj) {
  try {
    requireAdminToken_(token);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(PROVIDERS_SHEET_NAME);
    if (!sh) throw 'Providers sheet not found';
    
    sh.getRange(row, 1, 1, 17).setValues([[
      obj.name,
      obj.address,
      obj.phone,
      obj.monOpen,
      obj.monHours,
      obj.tueOpen,
      obj.tueHours,
      obj.wedOpen,
      obj.wedHours,
      obj.thuOpen,
      obj.thuHours,
      obj.friOpen,
      obj.friHours,
      obj.satOpen,
      obj.satHours,
      obj.sunOpen,
      obj.sunHours
    ]]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- DOCUMENT TYPES ---
function settingsGetDocumentTypes(token) {
  try {
    var list = genericList_(token, 'DocumentTypes', 2);
    var rows = list.map(function(item) { return { rowIdx: item.rowIdx, docType: item.row[0], updated: item.row[1] }; });
    return { success: true, rows: rows };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function settingsAddDocumentType(token, val) {
  try { return genericAdd_(token, 'DocumentTypes', [val]); } catch(e) { return { success: false, error: e.toString() }; }
}

function settingsDeleteDocumentType(token, row) { return genericDelete_(token, 'DocumentTypes', row); }

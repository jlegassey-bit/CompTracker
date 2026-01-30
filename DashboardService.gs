// --- DASHBOARD & ANALYTICS ---

// Constants
var SPREADSHEET_ID = "1RO1tTij8KY872eigB0Wm4tRqBExYKHj2p2SHd8WKy5c";
var FROIS_SHEET_NAME = "FROI Form";
var ADMINS_SHEET_NAME = "Admins";

function getDashboardStats(token) {
  try {
    requireAdminToken_(token);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(FROIS_SHEET_NAME);
    
    if (!sh || sh.getLastRow() < 2) {
      return { 
        success: true, 
        totalCases: 0,
        openCases: 0,
        closedCases: 0,
        avgDaysToClose: 0,
        byMonth: [],
        byDepartment: [],
        byType: []
      };
    }
    
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 33).getDisplayValues();
    
    var totalCases = data.length;
    var openCases = 0;
    var closedCases = 0;
    var daysToCloseArr = [];
    var byMonth = {};
    var byDepartment = {};
    var byType = {};
    
    var now = new Date();
    var sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
    
    data.forEach(function(row) {
      var incidentDate = row[10]; // Column K
      var department = row[4]; // Column E
      var injuryType = row[23]; // Column X - Nature of Injury
      var status = row[28] || 'New'; // Column AC - Status
      var submissionDate = row[27]; // Column AB
      var dateClosed = row[30]; // Column AE
      
      // Count by status
      if (status === 'Closed') {
        closedCases++;
        
        // Calculate days to close
        if (submissionDate && dateClosed) {
          try {
            var subDate = new Date(submissionDate);
            var closeDate = new Date(dateClosed);
            var days = Math.floor((closeDate - subDate) / (1000 * 60 * 60 * 24));
            if (days >= 0) daysToCloseArr.push(days);
          } catch (e) {}
        }
      } else {
        openCases++;
      }
      
      // Count by month (last 6 months)
      if (incidentDate) {
        try {
          var incDate = new Date(incidentDate);
          if (incDate >= sixMonthsAgo) {
            var monthKey = incDate.getFullYear() + '-' + String(incDate.getMonth() + 1).padStart(2, '0');
            byMonth[monthKey] = (byMonth[monthKey] || 0) + 1;
          }
        } catch (e) {}
      }
      
      // Count by department
      if (department) {
        byDepartment[department] = (byDepartment[department] || 0) + 1;
      }
      
      // Count by injury type
      if (injuryType) {
        byType[injuryType] = (byType[injuryType] || 0) + 1;
      }
    });
    
    // Calculate average days to close
    var avgDaysToClose = 0;
    if (daysToCloseArr.length > 0) {
      avgDaysToClose = Math.round(daysToCloseArr.reduce(function(a, b) { return a + b; }, 0) / daysToCloseArr.length);
    }
    
    // Format month data for chart (last 6 months)
    var monthLabels = [];
    var monthCounts = [];
    for (var i = 5; i >= 0; i--) {
      var d = new Date();
      d.setMonth(d.getMonth() - i);
      var key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      var monthName = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getMonth()];
      monthLabels.push(monthName + ' ' + d.getFullYear());
      monthCounts.push(byMonth[key] || 0);
    }
    
    // Top 5 departments
    var deptArray = Object.keys(byDepartment).map(function(k) {
      return { department: k, count: byDepartment[k] };
    });
    deptArray.sort(function(a, b) { return b.count - a.count; });
    var top5Depts = deptArray.slice(0, 5);
    
    // Top 5 injury types
    var typeArray = Object.keys(byType).map(function(k) {
      return { type: k, count: byType[k] };
    });
    typeArray.sort(function(a, b) { return b.count - a.count; });
    var top5Types = typeArray.slice(0, 5);
    
    return {
      success: true,
      totalCases: totalCases,
      openCases: openCases,
      closedCases: closedCases,
      avgDaysToClose: avgDaysToClose,
      monthLabels: monthLabels,
      monthCounts: monthCounts,
      top5Depts: top5Depts,
      top5Types: top5Types
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateCaseStatus(token, rowIdx, status, assignedTo) {
  try {
    var userEmail = requireAdminToken_(token);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(FROIS_SHEET_NAME);
    
    if (!sh || rowIdx < 2 || rowIdx > sh.getLastRow()) {
      return { success: false, error: 'Invalid row' };
    }
    
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
    
    // Update Status (Column AC = 29)
    if (status) {
      sh.getRange(rowIdx, 29).setValue(status);
      
      // If closing, set Date Closed (Column AE = 31)
      if (status === 'Closed') {
        sh.getRange(rowIdx, 31).setValue(now);
      }
    }
    
    // Update Assigned To (Column AD = 30)
    if (assignedTo !== undefined) {
      sh.getRange(rowIdx, 30).setValue(assignedTo);
    }
    
    // Update timestamp (Column AF = 32)
    sh.getRange(rowIdx, 32).setValue(now);
    
    // Log history
    logCaseHistory(rowIdx, userEmail, 'Status changed to: ' + status + (assignedTo ? ', Assigned to: ' + assignedTo : ''));
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getCaseHistory(token, rowIdx) {
  try {
    requireAdminToken_(token);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var histSh = ss.getSheetByName('Case_History');
    
    if (!histSh) {
      // Create history sheet if it doesn't exist
      histSh = ss.insertSheet('Case_History');
      histSh.getRange(1, 1, 1, 5).setValues([['Row Index', 'Timestamp', 'User', 'Action', 'Details']]);
      histSh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
      histSh.setFrozenRows(1);
    }
    
    if (histSh.getLastRow() < 2) {
      return { success: true, history: [] };
    }
    
    var data = histSh.getRange(2, 1, histSh.getLastRow() - 1, 5).getValues();
    var history = [];
    
    data.forEach(function(row) {
      if (row[0] == rowIdx) {
        history.push({
          timestamp: row[1],
          user: row[2],
          action: row[3],
          details: row[4]
        });
      }
    });
    
    // Sort by timestamp descending
    history.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    
    return { success: true, history: history };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function logCaseHistory(rowIdx, userEmail, details) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var histSh = ss.getSheetByName('Case_History');
    
    if (!histSh) {
      histSh = ss.insertSheet('Case_History');
      histSh.getRange(1, 1, 1, 5).setValues([['Row Index', 'Timestamp', 'User', 'Action', 'Details']]);
      histSh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
      histSh.setFrozenRows(1);
    }
    
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
    histSh.appendRow([rowIdx, now, userEmail, 'Update', details]);
    
  } catch (e) {
    Logger.log('Error logging history: ' + e.toString());
  }
}

function getAdminsList(token) {
  try {
    requireAdminToken_(token);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(ADMINS_SHEET_NAME);
    
    if (!sh || sh.getLastRow() < 2) {
      return { success: true, admins: [] };
    }
    
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
    var admins = [];
    
    data.forEach(function(row) {
      if (row[0] && row[1]) {
        admins.push({ name: row[0], email: row[1] });
      }
    });
    
    return { success: true, admins: admins };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateCaseAssignment(token, rowIdx, assignedTo) {
  try {
    var userEmail = requireAdminToken_(token);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(FROIS_SHEET_NAME);
    
    if (!sh || rowIdx < 2 || rowIdx > sh.getLastRow()) {
      return { success: false, error: 'Invalid row' };
    }
    
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
    
    // Update Assigned To (Column AD = 30)
    sh.getRange(rowIdx, 30).setValue(assignedTo);
    
    // Update timestamp (Column AF = 32)
    sh.getRange(rowIdx, 32).setValue(now);
    
    // Log history
    var assignedName = assignedTo || 'Unassigned';
    logCaseHistory(rowIdx, userEmail, 'Assigned to: ' + assignedName);
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
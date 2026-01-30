// --- ADMIN AUTH & TOKENS ---
var ADMIN_OTP_TTL_SECONDS = 10 * 60; // 10 minutes
var ADMIN_TOKEN_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours
var ADMIN_TOKEN_SECRET_KEY = 'ADMIN_TOKEN_SECRET_V1';

// --- NOTICE LINKS CONFIG ---
var NOTICE_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days
var NOTICE_SECRET_KEY = 'NOTICE_SECRET_V1';

// --- ADMIN HELPERS ---

function requireAdmin_(email) {
  var em = normalizeEmail_(email);
  if (!isAllowedDomain_(em)) throw new Error('Unauthorized: invalid domain.');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) throw new Error('Unauthorized: admin table not configured.');

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 4)).getDisplayValues();
  for (var i = 0; i < data.length; i++) {
    if (normalizeEmail_(data[i][0]) === em) return true;
  }
  throw new Error('Unauthorized.');
}

function getAdminTokenSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty(ADMIN_TOKEN_SECRET_KEY);
  if (!secret) {
    secret = Utilities.getUuid();
    props.setProperty(ADMIN_TOKEN_SECRET_KEY, secret);
  }
  return secret;
}

function getNoticeSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty(NOTICE_SECRET_KEY);
  if (!secret) {
    secret = Utilities.getUuid();
    props.setProperty(NOTICE_SECRET_KEY, secret);
  }
  return secret;
}

function requireAdminToken_(token) {
  if (!token) throw new Error('Unauthorized: missing token');
  var parsed = parseAdminToken_(token);
  if (!parsed || !parsed.valid) throw new Error('Unauthorized: invalid token');
  requireAdmin_(parsed.email);
  return parsed.email;
}

// --- OTP (LOGIN) ---

function adminSendOtp(email) {
  try {
    var em = normalizeEmail_(email);
    requireAdmin_(em);
    var otp = generateOtp_();
    var expiry = Date.now() + (ADMIN_OTP_TTL_SECONDS * 1000);
    var cache = CacheService.getScriptCache();
    var key = 'otp_' + em;
    cache.put(key, JSON.stringify({ otp: otp, exp: expiry }), ADMIN_OTP_TTL_SECONDS);

    var subject = 'FROI Admin Login Code';
    var body = 'Your login code: ' + otp + '\n\nExpires in ' + (ADMIN_OTP_TTL_SECONDS / 60) + ' minutes.';
    MailApp.sendEmail(em, subject, body);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function adminVerifyOtp(email, otp) {
  try {
    var em = normalizeEmail_(email);
    requireAdmin_(em);

    var cache = CacheService.getScriptCache();
    var key = 'otp_' + em;
    var stored = cache.get(key);
    if (!stored) return { success: false, error: 'OTP expired or not found' };

    var obj = JSON.parse(stored);
    if (obj.otp !== otp) return { success: false, error: 'Invalid OTP' };
    if (Date.now() > obj.exp) return { success: false, error: 'OTP expired' };

    cache.remove(key);
    
    // Get admin details from Admins sheet
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(ADMIN_SHEET_NAME);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues();
    
    var userName = '';
    var userTitle = '';
    var userDept = '';
    
    for (var i = 0; i < data.length; i++) {
      if (normalizeEmail_(data[i][0]) === em) { // Column A is email
        userName = data[i][1];  // Column B is name
        userTitle = data[i][2]; // Column C is title
        userDept = data[i][3];  // Column D is department
        break;
      }
    }
    
    var token = generateAdminToken_(em);
    return { 
      success: true, 
      token: token, 
      user: { 
        email: em,
        name: userName,
        title: userTitle,
        department: userDept
      }
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function generateOtp_() {
  var code = '';
  for (var i = 0; i < 6; i++) {
    code += Math.floor(Math.random() * 10).toString();
  }
  return code;
}

// --- ADMIN TOKEN ---

function generateAdminToken_(email) {
  var expiry = Date.now() + ADMIN_TOKEN_TTL_MS;
  var payload = { email: email, exp: expiry };
  var payloadStr = JSON.stringify(payload);
  var secret = getAdminTokenSecret_();
  var signature = Utilities.computeHmacSha256Signature(payloadStr, secret);
  var sigB64 = Utilities.base64Encode(signature);
  var tokenObj = { payload: payload, sig: sigB64 };
  return JSON.stringify(tokenObj);
}

function parseAdminToken_(token) {
  try {
    var obj = JSON.parse(token);
    var payloadStr = JSON.stringify(obj.payload);
    var secret = getAdminTokenSecret_();
    var expectedSig = Utilities.computeHmacSha256Signature(payloadStr, secret);
    var expectedB64 = Utilities.base64Encode(expectedSig);

    if (obj.sig !== expectedB64) return { valid: false };
    if (Date.now() > obj.payload.exp) return { valid: false };

    return { valid: true, email: obj.payload.email };
  } catch (e) {
    return { valid: false };
  }
}

// --- NOTICE SIGNATURE HELPERS ---

function computeNoticeSig_(recordId, rowIdx, ts) {
  var secret = getNoticeSecret_();
  var data = recordId + '|' + rowIdx + '|' + ts;
  var sig = Utilities.computeHmacSha256Signature(data, secret);
  return Utilities.base64EncodeWebSafe(sig);
}

function verifyNoticeSig_(recordId, rowIdx, ts, sig) {
  if (!recordId || !rowIdx || !ts || !sig) {
    return { ok: false, reason: 'missing' };
  }
  var nowMs = Date.now();
  var linkMs = parseInt(ts, 10);
  if (isNaN(linkMs) || (nowMs - linkMs) > NOTICE_TTL_MS) {
    return { ok: false, reason: 'expired' };
  }
  var expected = computeNoticeSig_(recordId, rowIdx, ts);
  if (sig !== expected) {
    return { ok: false, reason: 'invalid' };
  }
  return { ok: true };
}

// --- PUBLIC FACING ---

function adminInitiateLogin(email) {
  return adminSendOtp(email);
}

function adminLogin(email, otp) {
  return adminVerifyOtp(email, otp);
}

function adminCheckToken(token) {
  try {
    var parsed = parseAdminToken_(token);
    if (!parsed || !parsed.valid) {
      return { valid: false };
    }
    requireAdmin_(parsed.email);
    return { valid: true, email: parsed.email };
  } catch (e) {
    return { valid: false };
  }
}

function adminLogout(token) {
  return { success: true };
}

function getNoticeData(recordId, rowIdx) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var resSheet = ss.getSheetByName(RESTR_SHEET_NAME);
  var r = resSheet.getRange(rowIdx, 1, 1, 12).getDisplayValues()[0];
  
  var froiSheet = ss.getSheetByName(FORM_SHEET_NAME);
  // Column 9 is Employee Name (with Record ID in column 1)
  var empName = froiSheet.getRange(recordId, 9).getValue();

  return {
    employee: empName,
    provider: r[3],
    apptDate: r[4],
    apptTime: r[5],
    restrictions: r[6],
    followDate: r[7],
    followTime: r[8],
    followProv: r[9],
    adminName: r[10],
    adminTitle: r[11],
    entryDate: r[1]
  };
}

// --- FRONTEND WRAPPER FUNCTIONS ---
// These functions match what the frontend HTML expects

function requestAuthCode(email) {
  return adminInitiateLogin(email);
}

function verifyAndGetData(email, code) {
  try {
    var loginResult = adminLogin(email, code);
    
    if (!loginResult.success) {
      return loginResult;
    }
    
    // Get the cases/reports data
    var casesResult = adminGetCases(loginResult.token);
    
    if (!casesResult.success) {
      return casesResult;
    }
    
    // Return combined result
   return {
      success: true,
      token: loginResult.token,
      user: loginResult.user,  // ← Changed from loginResult.email to loginResult.user
      data: casesResult.reports
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
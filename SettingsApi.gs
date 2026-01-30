/**
 * SettingsApi.gs
 * Server-side functions for Modal_Settings.html
 * Fixes: google.script.run....listAdmins is not a function
 *
 * Storage sheets (auto-created if missing):
 * - Admins
 * - Departments
 * - WorkLocations
 * - PrimaryCauses
 * - Treatments
 * - TreatmentProviders
 * - BodyParts
 */

function listAdmins(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Admins', ['Email', 'Name', 'Title', 'Department', 'Updated']);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      email: (r.Email || '').trim(),
      name: (r.Name || '').trim(),
      title: (r.Title || '').trim(),
      department: (r.Department || '').trim(),
      updated: (r.Updated || '').trim()
    })).filter(r => r.email);
    return { success: true, rows: out };
  });
}

function addAdmin(requestingEmail, payload) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Admins', ['Email', 'Name', 'Title', 'Department', 'Updated']);
    const email = ((payload && payload.email) || '').trim();
    if (!email) return { success: false, error: 'Email is required.' };

    // prevent duplicates (case-insensitive)
    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r.Email || '').trim().toLowerCase() === email.toLowerCase());
    if (exists) return { success: false, error: 'That email already has admin access.' };

    sh.appendRow([
      email,
      ((payload && payload.name) || '').trim(),
      ((payload && payload.title) || '').trim(),
      ((payload && payload.department) || '').trim(),
      nowStamp_()
    ]);
    return { success: true };
  });
}

function updateAdmin(requestingEmail, rowIdx, payload) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Admins', ['Email', 'Name', 'Title', 'Department', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const email = ((payload && payload.email) || '').trim();
    if (!email) return { success: false, error: 'Email is required.' };

    // prevent duplicates against other rows
    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj.Email || '').trim().toLowerCase() === email.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another admin row already uses that email.' };

    writeRowByHeaders_(sh, r, {
      Email: email,
      Name: ((payload && payload.name) || '').trim(),
      Title: ((payload && payload.title) || '').trim(),
      Department: ((payload && payload.department) || '').trim(),
      Updated: nowStamp_()
    });
    return { success: true };
  });
}

function deleteAdmin(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Admins', ['Email', 'Name', 'Title', 'Department', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ---------------- Departments ----------------

function listDepartments(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Departments', [
      'Department',
      'HR Name','HR Email','HR Phone',
      'Payroll Name','Payroll Email','Payroll Phone',
      'Safety Name','Safety Email','Safety Phone',
      'Updated'
    ]);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      department: (r['Department'] || '').trim(),
      hrName: (r['HR Name'] || '').trim(),
      hrEmail: (r['HR Email'] || '').trim(),
      hrPhone: (r['HR Phone'] || '').trim(),
      payrollName: (r['Payroll Name'] || '').trim(),
      payrollEmail: (r['Payroll Email'] || '').trim(),
      payrollPhone: (r['Payroll Phone'] || '').trim(),
      safetyName: (r['Safety Name'] || '').trim(),
      safetyEmail: (r['Safety Email'] || '').trim(),
      safetyPhone: (r['Safety Phone'] || '').trim(),
      updated: (r['Updated'] || '').trim()
    })).filter(r => r.department);
    return { success: true, rows: out };
  });
}

function addDepartment(requestingEmail, payload) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Departments', [
      'Department',
      'HR Name','HR Email','HR Phone',
      'Payroll Name','Payroll Email','Payroll Phone',
      'Safety Name','Safety Email','Safety Phone',
      'Updated'
    ]);
    const dept = ((payload && payload.department) || '').trim();
    if (!dept) return { success: false, error: 'Department name is required.' };

    // prevent duplicates
    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r['Department'] || '').trim().toLowerCase() === dept.toLowerCase());
    if (exists) return { success: false, error: 'That department already exists.' };

    sh.appendRow([
      dept,
      ((payload && payload.hrName) || '').trim(),
      ((payload && payload.hrEmail) || '').trim(),
      ((payload && payload.hrPhone) || '').trim(),
      ((payload && payload.payrollName) || '').trim(),
      ((payload && payload.payrollEmail) || '').trim(),
      ((payload && payload.payrollPhone) || '').trim(),
      ((payload && payload.safetyName) || '').trim(),
      ((payload && payload.safetyEmail) || '').trim(),
      ((payload && payload.safetyPhone) || '').trim(),
      nowStamp_()
    ]);
    return { success: true };
  });
}

function updateDepartment(requestingEmail, rowIdx, payload) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Departments', [
      'Department',
      'HR Name','HR Email','HR Phone',
      'Payroll Name','Payroll Email','Payroll Phone',
      'Safety Name','Safety Email','Safety Phone',
      'Updated'
    ]);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const dept = ((payload && payload.department) || '').trim();
    if (!dept) return { success: false, error: 'Department name is required.' };

    // prevent duplicates against other rows
    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj['Department'] || '').trim().toLowerCase() === dept.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another row already uses that department name.' };

    writeRowByHeaders_(sh, r, {
      'Department': dept,
      'HR Name': ((payload && payload.hrName) || '').trim(),
      'HR Email': ((payload && payload.hrEmail) || '').trim(),
      'HR Phone': ((payload && payload.hrPhone) || '').trim(),
      'Payroll Name': ((payload && payload.payrollName) || '').trim(),
      'Payroll Email': ((payload && payload.payrollEmail) || '').trim(),
      'Payroll Phone': ((payload && payload.payrollPhone) || '').trim(),
      'Safety Name': ((payload && payload.safetyName) || '').trim(),
      'Safety Email': ((payload && payload.safetyEmail) || '').trim(),
      'Safety Phone': ((payload && payload.safetyPhone) || '').trim(),
      'Updated': nowStamp_()
    });
    return { success: true };
  });
}

function deleteDepartment(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Departments', [
      'Department',
      'HR Name','HR Email','HR Phone',
      'Payroll Name','Payroll Email','Payroll Phone',
      'Safety Name','Safety Email','Safety Phone',
      'Updated'
    ]);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ---------------- Work Locations ----------------

function listWorkLocations(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('WorkLocations', ['Work Location', 'Department', 'Updated']);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      workLocation: (r['Work Location'] || '').trim(),
      department: (r['Department'] || '').trim(),
      updated: (r['Updated'] || '').trim()
    })).filter(r => r.workLocation);
    return { success: true, rows: out };
  });
}

function addWorkLocation(requestingEmail, workLocation, department) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('WorkLocations', ['Work Location', 'Department', 'Updated']);
    const wl = String(workLocation || '').trim();
    const dept = String(department || '').trim();
    if (!wl) return { success: false, error: 'Work Location is required.' };
    if (!dept) return { success: false, error: 'Department is required.' };

    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r['Work Location'] || '').trim().toLowerCase() === wl.toLowerCase());
    if (exists) return { success: false, error: 'That work location already exists.' };

    sh.appendRow([wl, dept, nowStamp_()]);
    return { success: true };
  });
}

function updateWorkLocation(requestingEmail, rowIdx, workLocation, department) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('WorkLocations', ['Work Location', 'Department', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const wl = String(workLocation || '').trim();
    const dept = String(department || '').trim();
    if (!wl) return { success: false, error: 'Work Location is required.' };
    if (!dept) return { success: false, error: 'Department is required.' };

    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj['Work Location'] || '').trim().toLowerCase() === wl.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another row already uses that work location.' };

    writeRowByHeaders_(sh, r, {
      'Work Location': wl,
      'Department': dept,
      'Updated': nowStamp_()
    });
    return { success: true };
  });
}

function deleteWorkLocation(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('WorkLocations', ['Work Location', 'Department', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ---------------- Primary Causes ----------------

function listPrimaryCauses(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('PrimaryCauses', ['Primary Cause', 'Updated']);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      cause: (r['Primary Cause'] || '').trim(),
      updated: (r['Updated'] || '').trim()
    })).filter(r => r.cause);
    return { success: true, rows: out };
  });
}

function addPrimaryCause(requestingEmail, cause) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('PrimaryCauses', ['Primary Cause', 'Updated']);
    const c = String(cause || '').trim();
    if (!c) return { success: false, error: 'Primary Cause is required.' };

    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r['Primary Cause'] || '').trim().toLowerCase() === c.toLowerCase());
    if (exists) return { success: false, error: 'That primary cause already exists.' };

    sh.appendRow([c, nowStamp_()]);
    return { success: true };
  });
}

function updatePrimaryCause(requestingEmail, rowIdx, cause) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('PrimaryCauses', ['Primary Cause', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const c = String(cause || '').trim();
    if (!c) return { success: false, error: 'Primary Cause is required.' };

    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj['Primary Cause'] || '').trim().toLowerCase() === c.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another row already uses that primary cause.' };

    writeRowByHeaders_(sh, r, { 'Primary Cause': c, 'Updated': nowStamp_() });
    return { success: true };
  });
}

function deletePrimaryCause(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('PrimaryCauses', ['Primary Cause', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ---------------- Treatments ----------------

function listTreatments(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Treatments', ['Treatment', 'Updated']);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      treatment: (r['Treatment'] || '').trim(),
      updated: (r['Updated'] || '').trim()
    })).filter(r => r.treatment);
    return { success: true, rows: out };
  });
}

function addTreatment(requestingEmail, treatment) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Treatments', ['Treatment', 'Updated']);
    const t = String(treatment || '').trim();
    if (!t) return { success: false, error: 'Treatment is required.' };

    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r['Treatment'] || '').trim().toLowerCase() === t.toLowerCase());
    if (exists) return { success: false, error: 'That treatment already exists.' };

    sh.appendRow([t, nowStamp_()]);
    return { success: true };
  });
}

function updateTreatment(requestingEmail, rowIdx, treatment) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Treatments', ['Treatment', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const t = String(treatment || '').trim();
    if (!t) return { success: false, error: 'Treatment is required.' };

    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj['Treatment'] || '').trim().toLowerCase() === t.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another row already uses that treatment.' };

    writeRowByHeaders_(sh, r, { 'Treatment': t, 'Updated': nowStamp_() });
    return { success: true };
  });
}

function deleteTreatment(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('Treatments', ['Treatment', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ---------------- Treatment Providers ----------------

function listTreatmentProviders(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('TreatmentProviders', [
      'Provider Name','Address','Phone',
      'Mon Open','Mon Hours',
      'Tue Open','Tue Hours',
      'Wed Open','Wed Hours',
      'Thu Open','Thu Hours',
      'Fri Open','Fri Hours',
      'Sat Open','Sat Hours',
      'Sun Open','Sun Hours',
      'Updated'
    ]);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      providerName: (r['Provider Name'] || '').trim(),
      address: (r['Address'] || '').trim(),
      phone: (r['Phone'] || '').trim(),
      monOpen: truthy_(r['Mon Open']),
      monHours: (r['Mon Hours'] || '').trim(),
      tueOpen: truthy_(r['Tue Open']),
      tueHours: (r['Tue Hours'] || '').trim(),
      wedOpen: truthy_(r['Wed Open']),
      wedHours: (r['Wed Hours'] || '').trim(),
      thuOpen: truthy_(r['Thu Open']),
      thuHours: (r['Thu Hours'] || '').trim(),
      friOpen: truthy_(r['Fri Open']),
      friHours: (r['Fri Hours'] || '').trim(),
      satOpen: truthy_(r['Sat Open']),
      satHours: (r['Sat Hours'] || '').trim(),
      sunOpen: truthy_(r['Sun Open']),
      sunHours: (r['Sun Hours'] || '').trim(),
      updated: (r['Updated'] || '').trim()
    })).filter(r => r.providerName);
    return { success: true, rows: out };
  });
}

function addTreatmentProvider(requestingEmail, payload) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('TreatmentProviders', [
      'Provider Name','Address','Phone',
      'Mon Open','Mon Hours',
      'Tue Open','Tue Hours',
      'Wed Open','Wed Hours',
      'Thu Open','Thu Hours',
      'Fri Open','Fri Hours',
      'Sat Open','Sat Hours',
      'Sun Open','Sun Hours',
      'Updated'
    ]);
    const name = ((payload && payload.providerName) || '').trim();
    if (!name) return { success: false, error: 'Provider Name is required.' };

    // prevent duplicates
    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r['Provider Name'] || '').trim().toLowerCase() === name.toLowerCase());
    if (exists) return { success: false, error: 'That provider already exists.' };

    sh.appendRow([
      name,
      ((payload && payload.address) || '').trim(),
      ((payload && payload.phone) || '').trim(),
      boolToCell_(payload && payload.monOpen), ((payload && payload.monHours) || '').trim(),
      boolToCell_(payload && payload.tueOpen), ((payload && payload.tueHours) || '').trim(),
      boolToCell_(payload && payload.wedOpen), ((payload && payload.wedHours) || '').trim(),
      boolToCell_(payload && payload.thuOpen), ((payload && payload.thuHours) || '').trim(),
      boolToCell_(payload && payload.friOpen), ((payload && payload.friHours) || '').trim(),
      boolToCell_(payload && payload.satOpen), ((payload && payload.satHours) || '').trim(),
      boolToCell_(payload && payload.sunOpen), ((payload && payload.sunHours) || '').trim(),
      nowStamp_()
    ]);
    return { success: true };
  });
}

function updateTreatmentProvider(requestingEmail, rowIdx, payload) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('TreatmentProviders', [
      'Provider Name','Address','Phone',
      'Mon Open','Mon Hours',
      'Tue Open','Tue Hours',
      'Wed Open','Wed Hours',
      'Thu Open','Thu Hours',
      'Fri Open','Fri Hours',
      'Sat Open','Sat Hours',
      'Sun Open','Sun Hours',
      'Updated'
    ]);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const name = ((payload && payload.providerName) || '').trim();
    if (!name) return { success: false, error: 'Provider Name is required.' };

    // prevent duplicates against other rows
    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj['Provider Name'] || '').trim().toLowerCase() === name.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another row already uses that provider name.' };

    writeRowByHeaders_(sh, r, {
      'Provider Name': name,
      'Address': ((payload && payload.address) || '').trim(),
      'Phone': ((payload && payload.phone) || '').trim(),
      'Mon Open': boolToCell_(payload && payload.monOpen),
      'Mon Hours': ((payload && payload.monHours) || '').trim(),
      'Tue Open': boolToCell_(payload && payload.tueOpen),
      'Tue Hours': ((payload && payload.tueHours) || '').trim(),
      'Wed Open': boolToCell_(payload && payload.wedOpen),
      'Wed Hours': ((payload && payload.wedHours) || '').trim(),
      'Thu Open': boolToCell_(payload && payload.thuOpen),
      'Thu Hours': ((payload && payload.thuHours) || '').trim(),
      'Fri Open': boolToCell_(payload && payload.friOpen),
      'Fri Hours': ((payload && payload.friHours) || '').trim(),
      'Sat Open': boolToCell_(payload && payload.satOpen),
      'Sat Hours': ((payload && payload.satHours) || '').trim(),
      'Sun Open': boolToCell_(payload && payload.sunOpen),
      'Sun Hours': ((payload && payload.sunHours) || '').trim(),
      'Updated': nowStamp_()
    });
    return { success: true };
  });
}

function deleteTreatmentProvider(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('TreatmentProviders', [
      'Provider Name','Address','Phone',
      'Mon Open','Mon Hours',
      'Tue Open','Tue Hours',
      'Wed Open','Wed Hours',
      'Thu Open','Thu Hours',
      'Fri Open','Fri Hours',
      'Sat Open','Sat Hours',
      'Sun Open','Sun Hours',
      'Updated'
    ]);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ---------------- Body Parts ----------------

function listBodyParts(requestingEmail) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('BodyParts', ['Body Part', 'Updated']);
    const rows = readSheetObjects_(sh);
    const out = rows.map(r => ({
      rowIdx: r.__rowIdx,
      bodyPart: (r['Body Part'] || '').trim(),
      updated: (r['Updated'] || '').trim()
    })).filter(r => r.bodyPart);
    return { success: true, rows: out };
  });
}

function addBodyPart(requestingEmail, bodyPart) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('BodyParts', ['Body Part', 'Updated']);
    const bp = String(bodyPart || '').trim();
    if (!bp) return { success: false, error: 'Body Part is required.' };

    const rows = readSheetObjects_(sh);
    const exists = rows.some(r => String(r['Body Part'] || '').trim().toLowerCase() === bp.toLowerCase());
    if (exists) return { success: false, error: 'That body part already exists.' };

    sh.appendRow([bp, nowStamp_()]);
    return { success: true };
  });
}

function updateBodyPart(requestingEmail, rowIdx, bodyPart) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('BodyParts', ['Body Part', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };

    const bp = String(bodyPart || '').trim();
    if (!bp) return { success: false, error: 'Body Part is required.' };

    const rows = readSheetObjects_(sh);
    const dup = rows.some(obj =>
      obj.__rowIdx !== r &&
      String(obj['Body Part'] || '').trim().toLowerCase() === bp.toLowerCase()
    );
    if (dup) return { success: false, error: 'Another row already uses that body part.' };

    writeRowByHeaders_(sh, r, { 'Body Part': bp, 'Updated': nowStamp_() });
    return { success: true };
  });
}

function deleteBodyPart(requestingEmail, rowIdx) {
  return withAdminGuard_(requestingEmail, function () {
    const sh = getOrCreateSheet_('BodyParts', ['Body Part', 'Updated']);
    const r = Number(rowIdx);
    if (!r || r < 2) return { success: false, error: 'Invalid row index.' };
    sh.deleteRow(r);
    return { success: true };
  });
}

// ================= Helpers =================

function withAdminGuard_(email, fn) {
  try {
    const e = String(email || '').trim();
    if (!e) return { success: false, error: 'Missing user identity.' };
    if (!isAdmin_(e)) return { success: false, error: 'Access denied (admin only).' };
    return fn();
  } catch (err) {
    return { success: false, error: String(err && err.message ? err.message : err) };
  }
}

function isAdmin_(email) {
  const sh = getOrCreateSheet_('Admins', ['Email', 'Name', 'Title', 'Department', 'Updated']);
  const rows = readSheetObjects_(sh);
  const target = String(email || '').trim().toLowerCase();
  return rows.some(r => String(r.Email || '').trim().toLowerCase() === target);
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  // ensure headers
  const lastCol = sh.getLastColumn();
  const firstRow = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const hasAnyHeader = firstRow.some(v => String(v || '').trim() !== '');
  if (!hasAnyHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readSheetObjects_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const rowObj = { __rowIdx: i + 2 };
    for (let c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      rowObj[headers[c]] = values[i][c];
    }
    out.push(rowObj);
  }
  return out;
}

function writeRowByHeaders_(sh, rowIdx, obj) {
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const row = sh.getRange(rowIdx, 1, 1, lastCol).getValues()[0];

  Object.keys(obj || {}).forEach(function (k) {
    const key = String(k || '').trim();
    if (!key) return;
    const colIndex = headers.indexOf(key);
    if (colIndex === -1) return; // ignore unknown headers
    row[colIndex] = obj[key];
  });

  sh.getRange(rowIdx, 1, 1, lastCol).setValues([row]);
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd HH:mm:ss');
}

function truthy_(v) {
  if (v === true) return true;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

function boolToCell_(b) {
  return b ? true : false;
}

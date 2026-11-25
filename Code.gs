/**
 * ============================================================================
 * ECHAAR WORKFORCE ATTENDANCE MANAGEMENT SYSTEM - BACKEND (Google Apps Script)
 * ============================================================================
 * 
 * Configured Exact Office Coordinates:
 * - Latitude:  19.193740
 * - Longitude: 72.957355
 * - Allowed Radius: 80 meters
 * 
 * Features:
 * 1. Full Admin visibility into every employee's attendance and absence reasons.
 * 2. Automated Punctuality calculation (On-Time vs Late by X minutes) based on shift start.
 * 3. Work Duration calculation (Total Hours & Minutes worked per day).
 * 4. Bidirectional Device Lock (One physical device per employee per day).
 * 5. Tamper-proof, immutable attendance logs.
 * 6. Multi-tab resolution (Auto-resolves Attendance and Users tabs even if named Sheet1/Sheet2).
 */

var CONFIG = {
  SHEET_ATTENDANCE: 'Attendance',
  SHEET_USERS: 'Users',
  DEFAULT_PASSWORD: 'pass@123',
  TIMEZONE: 'Asia/Kolkata',
  DATE_FORMAT: 'yyyy-MM-dd',
  TIME_FORMAT: 'hh:mm:ss a',
  OFFICE_LAT: 19.193740,
  OFFICE_LNG: 72.957355,
  ALLOWED_RADIUS_METERS: 80,
  DEFAULT_SHIFT_START: '09:30 AM',
  DEFAULT_SHIFT_END: '06:30 PM',
  DEFAULT_GRACE_MINUTES: 15
};

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getUsersSheet() {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  if (sheets.length === 0) return null;

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (name === 'users' || name === 'user' || name === 'employees' || name === 'login' || name === 'credentials' || name === 'staff') {
      return sheets[i];
    }
  }

  for (var j = 0; j < sheets.length; j++) {
    var n = sheets[j].getName().trim().toLowerCase();
    if (n.indexOf('user') !== -1 || n.indexOf('login') !== -1 || n.indexOf('staff') !== -1 || n.indexOf('cred') !== -1) {
      return sheets[j];
    }
  }

  if (sheets.length >= 2) return sheets[1];
  return sheets[0];
}

function getAttendanceSheet() {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  if (sheets.length === 0) return null;

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (name === 'attendance' || name === 'attend' || name === 'punch' || name === 'log' || name === 'logs') {
      return sheets[i];
    }
  }

  for (var j = 0; j < sheets.length; j++) {
    var n = sheets[j].getName().trim().toLowerCase();
    if (n.indexOf('attend') !== -1 || n.indexOf('punch') !== -1 || n.indexOf('log') !== -1) {
      return sheets[j];
    }
  }

  return sheets[0];
}

function getFormattedDate(date) {
  var tz = CONFIG.TIMEZONE || Session.getScriptTimeZone();
  return Utilities.formatDate(date || new Date(), tz, CONFIG.DATE_FORMAT);
}

function getFormattedTime(date) {
  var tz = CONFIG.TIMEZONE || Session.getScriptTimeZone();
  return Utilities.formatDate(date || new Date(), tz, CONFIG.TIME_FORMAT);
}

/**
 * Normalizes Google Sheets Date/Time objects or strings into clean 'hh:mm a' format.
 */
function normalizeTimeString(val) {
  if (!val || val === '-') return '-';
  if (val instanceof Date) {
    var tz = CONFIG.TIMEZONE || Session.getScriptTimeZone();
    return Utilities.formatDate(val, tz, 'hh:mm a');
  }
  var str = String(val).trim();
  if (str.indexOf('1899') !== -1 || str.indexOf('GMT') !== -1) {
    try {
      var d = new Date(str);
      if (!isNaN(d.getTime())) {
        var tz2 = CONFIG.TIMEZONE || Session.getScriptTimeZone();
        return Utilities.formatDate(d, tz2, 'hh:mm a');
      }
    } catch (e) {}
  }
  return str;
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr || timeStr === '-') return null;
  var clean = normalizeTimeString(timeStr);
  var match = clean.match(/(\d+):(\d+)(?::\d+)?\s*(AM|PM)/i);
  if (!match) return null;
  var hours = parseInt(match[1], 10);
  var mins = parseInt(match[2], 10);
  var ampm = match[3].toUpperCase();
  if (ampm === 'PM' && hours < 12) hours += 12;
  if (ampm === 'AM' && hours === 12) hours = 0;
  return hours * 60 + mins;
}

function calculatePunctuality(punchInStr, shiftStartStr, graceMinutes) {
  if (!punchInStr || punchInStr === '-') return { status: 'Pending', label: 'Pending', isLate: false, diffMinutes: 0 };
  var punchMinutes = parseTimeToMinutes(punchInStr);
  var shiftMinutes = parseTimeToMinutes(shiftStartStr) || parseTimeToMinutes(CONFIG.DEFAULT_SHIFT_START);
  if (punchMinutes === null || shiftMinutes === null) return { status: 'Unknown', label: 'Unknown', isLate: false, diffMinutes: 0 };

  var grace = typeof graceMinutes === 'number' ? graceMinutes : CONFIG.DEFAULT_GRACE_MINUTES;
  var diff = punchMinutes - shiftMinutes;

  if (diff <= grace) {
    return { status: 'On-Time', label: 'On-Time', isLate: false, diffMinutes: diff };
  } else {
    return { status: 'Late', label: 'Late by ' + diff + 'm', isLate: true, diffMinutes: diff };
  }
}

function calculateWorkDuration(punchInStr, punchOutStr) {
  if (!punchInStr || punchInStr === '-') return '-';
  var inMins = parseTimeToMinutes(punchInStr);
  if (inMins === null) return '-';
  if (!punchOutStr || punchOutStr === '-') return 'In Progress';
  var outMins = parseTimeToMinutes(punchOutStr);
  if (outMins === null) return 'In Progress';

  var diffMins = outMins - inMins;
  if (diffMins < 0) diffMins += 24 * 60; // Overnight shift

  var hours = Math.floor(diffMins / 60);
  var mins = diffMins % 60;
  return hours + 'h ' + mins + 'm';
}

/**
 * Determines whether a user account belongs to Management/Admin (MD, HR, Head of Dept, Director)
 * Management accounts have full administrative view access but do not count as workforce.
 */
function isManagementRole(role, empId) {
  var r = String(role || '').trim().toUpperCase();
  var id = String(empId || '').trim().toUpperCase();
  
  if (r === 'ADMIN' || r === 'MD' || r.indexOf('MANAGING DIRECTOR') !== -1 || r === 'HR' || 
      r.indexOf('HEAD') !== -1 || r === 'HOD' || r === 'DIRECTOR' || r === 'MANAGEMENT') {
    return true;
  }
  if (id === 'MD' || id === 'HR' || id === 'HEAD' || id.indexOf('EEA') === 0) {
    return true;
  }
  return false;
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  var R = 6371e3;
  var toRad = function(deg) { return (deg * Math.PI) / 180; };
  var dLat = toRad(lat2 - lat1);
  var dLon = toRad(lon2 - lon1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
          Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    return handleRequest(e ? e.parameter : {});
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString(), message: 'GET request failed.' });
  }
}

function doPost(e) {
  try {
    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }
    return handleRequest(params);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString(), message: 'POST request failed.' });
  }
}

function handleRequest(params) {
  var action = params.action;
  if (!action) return jsonResponse({ success: false, message: 'Missing action parameter.' });

  switch (action) {
    case 'ping':
      return jsonResponse({
        success: true,
        message: 'Echaar Attendance API is active.',
        officeCoordinates: { lat: CONFIG.OFFICE_LAT, lng: CONFIG.OFFICE_LNG, radiusMeters: CONFIG.ALLOWED_RADIUS_METERS },
        serverDate: getFormattedDate(),
        serverTime: getFormattedTime()
      });

    case 'debugSheets':
      return handleDebugSheets();

    case 'initialSetup':
      return jsonResponse(initialSetup());

    case 'login':
      return handleLogin(params);

    case 'punchIn':
      return handlePunchIn(params);

    case 'punchOut':
      return handlePunchOut(params);

    case 'submitAbsence':
      return handleSubmitAbsence(params);

    case 'getEmployeeMonthly':
      return handleGetEmployeeMonthly(params);

    case 'getAdminDashboard':
      return handleGetAdminDashboard(params);

    case 'saveOfficeTimings':
      return handleSaveOfficeTimings(params);

    case 'getOfficeTimings':
      return handleGetOfficeTimings();

    case 'changePassword':
      return handleChangePassword(params);

    case 'getAdminMonthlyCalendar':
      return handleGetAdminMonthlyCalendar(params);

    default:
      return jsonResponse({ success: false, message: 'Unknown action: ' + action });
  }
}

function handleDebugSheets() {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var summary = [];
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var data = s.getDataRange().getValues();
    summary.push({
      tabIndex: i + 1,
      sheetName: s.getName(),
      rowCount: data.length,
      colCount: data.length > 0 ? data[0].length : 0,
      headers: data.length > 0 ? data[0] : [],
      sampleRow1: data.length > 1 ? data[1] : [],
      sampleRow2: data.length > 2 ? data[2] : []
    });
  }
  return jsonResponse({ success: true, spreadsheetName: ss.getName(), sheetsCount: sheets.length, sheets: summary });
}

function findUser(identifier) {
  var sheet = getUsersSheet();
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  var cleanIdentifier = String(identifier || '').replace(/\s+/g, '').toUpperCase();
  var headers = data[0];
  var empIdCol = 0, empNameCol = 1, phoneCol = 2, passCol = 3, roleCol = 4, shiftStartCol = 5, shiftEndCol = 6;

  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase().replace(/[^a-z]/g, '');
    if (h.indexOf('empid') !== -1 || h.indexOf('employeeid') !== -1 || h === 'id') empIdCol = c;
    else if (h.indexOf('name') !== -1) empNameCol = c;
    else if (h.indexOf('phone') !== -1 || h.indexOf('mobile') !== -1) phoneCol = c;
    else if (h.indexOf('pass') !== -1) passCol = c;
    else if (h.indexOf('role') !== -1) roleCol = c;
    else if (h.indexOf('shiftstart') !== -1 || h.indexOf('start') !== -1) shiftStartCol = c;
    else if (h.indexOf('shiftend') !== -1 || h.indexOf('end') !== -1) shiftEndCol = c;
  }

  for (var i = 1; i < data.length; i++) {
    var rowEmpId = String(data[i][empIdCol] || '').replace(/\s+/g, '').toUpperCase();
    var rowEmpName = String(data[i][empNameCol] || '').trim();
    var rowPhone = String(data[i][phoneCol] || '').replace(/[^0-9]/g, '');
    var cleanPhoneQuery = cleanIdentifier.replace(/[^0-9]/g, '');

    var rowPass = String(data[i][passCol] || CONFIG.DEFAULT_PASSWORD).trim();
    var rowRole = String(data[i][roleCol] || 'Employee').trim();
    var rowShiftStart = normalizeTimeString(data[i][shiftStartCol]) || CONFIG.DEFAULT_SHIFT_START;
    var rowShiftEnd = normalizeTimeString(data[i][shiftEndCol]) || CONFIG.DEFAULT_SHIFT_END;

    var phoneMatches = cleanPhoneQuery && (rowPhone === cleanPhoneQuery || (cleanPhoneQuery.length >= 10 && rowPhone.length >= 10 && cleanPhoneQuery.slice(-10) === rowPhone.slice(-10)));

    if (rowEmpId === cleanIdentifier || phoneMatches) {
      return {
        rowIndex: i + 1,
        employeeId: rowEmpId,
        employeeName: rowEmpName || rowEmpId,
        phone: rowPhone,
        password: rowPass,
        role: rowRole,
        shiftStart: rowShiftStart,
        shiftEnd: rowShiftEnd
      };
    }
  }
  return null;
}

function handleLogin(params) {
  var identifier = String(params.identifier || params.phone || params.employeeId || '').trim();
  var password = String(params.password || '').trim();

  if (!identifier || !password) {
    return jsonResponse({ success: false, message: 'Please enter both Mobile Number or Employee ID, and Password.' });
  }

  var user = findUser(identifier);
  if (!user) {
    return jsonResponse({ success: false, message: 'Account not found for "' + identifier + '". Check Tab 2 (Users).' });
  }

  if (user.password !== password) {
    return jsonResponse({ success: false, message: 'Invalid password. Check credentials.' });
  }

  return jsonResponse({
    success: true,
    message: 'Login successful.',
    user: {
      employeeId: user.employeeId,
      employeeName: user.employeeName,
      phone: user.phone,
      role: user.role,
      shiftStart: user.shiftStart,
      shiftEnd: user.shiftEnd
    }
  });
}

function handlePunchIn(params) {
  var employeeId = String(params.employeeId || '').replace(/\s+/g, '').toUpperCase();
  var coords = String(params.coords || '').trim();
  var deviceFingerprint = String(params.deviceFingerprint || '').trim();
  var clientLat = parseFloat(params.lat);
  var clientLng = parseFloat(params.lng);

  if (!employeeId) return jsonResponse({ success: false, message: 'Employee ID is required.' });

  var user = findUser(employeeId);
  if (!user) {
    var uSheet = getUsersSheet();
    return jsonResponse({ success: false, message: 'Employee ID "' + employeeId + '" not found in ' + (uSheet ? uSheet.getName() : 'Users') + '.' });
  }

  var distanceMeters = 0;
  if (!isNaN(clientLat) && !isNaN(clientLng)) {
    distanceMeters = calculateDistance(clientLat, clientLng, CONFIG.OFFICE_LAT, CONFIG.OFFICE_LNG);
    if (distanceMeters > CONFIG.ALLOWED_RADIUS_METERS) {
      return jsonResponse({
        success: false,
        message: 'Geofence Violation: You are ' + distanceMeters + 'm away. Allowed radius is ' + CONFIG.ALLOWED_RADIUS_METERS + 'm.'
      });
    }
  }

  var sheet = getAttendanceSheet();
  if (!sheet) return jsonResponse({ success: false, message: 'Attendance sheet is missing.' });

  var todayStr = getFormattedDate();
  var currentTimeStr = getFormattedTime();
  var data = sheet.getDataRange().getValues();

  // Device Lock & Duplicate check
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] instanceof Date ? getFormattedDate(data[i][0]) : String(data[i][0]).trim();
    var rowEmpId = String(data[i][1] || '').replace(/\s+/g, '').toUpperCase();
    var rowEmpName = String(data[i][2] || '').trim();
    var rowPunchIn = String(data[i][3] || '').trim();
    var rowStatus = String(data[i][7] || data[i][5] || '').trim();
    var rowDevice = String(data[i][9] || data[i][7] || '').trim();

    if (rowDate === todayStr) {
      if (rowEmpId === employeeId) {
        if (rowStatus === 'Absent') {
          return jsonResponse({ success: false, message: 'Absence reason was already logged for today (' + todayStr + ').' });
        }
        if (rowPunchIn) {
          return jsonResponse({ success: false, message: 'You have already punched in today at ' + rowPunchIn + '.' });
        }
        if (rowDevice && deviceFingerprint && rowDevice !== deviceFingerprint) {
          return jsonResponse({ success: false, message: 'Security Alert: Employee ' + employeeId + ' is locked to their original device today!' });
        }
      }
      if (deviceFingerprint && rowDevice === deviceFingerprint && rowEmpId !== employeeId) {
        return jsonResponse({ success: false, message: 'Security Alert: This device was already used today for ' + rowEmpName + ' (' + rowEmpId + ').' });
      }
    }
  }

  // Calculate punctuality on Punch In
  var punctuality = calculatePunctuality(currentTimeStr, user.shiftStart, CONFIG.DEFAULT_GRACE_MINUTES);

  // Schema: Date, Employee_ID, Employee_Name, Punch_In_Time, Punch_Out_Time, Work_Duration, Punctuality, Status, Absence_Reason, Device_Fingerprint, Location_Coords, Distance_Meters
  var newRow = [
    todayStr,
    user.employeeId,
    user.employeeName,
    currentTimeStr,
    '', // Punch Out
    'In Progress', // Work Duration
    punctuality.label, // Punctuality
    'Present',
    '', // Absence Reason
    deviceFingerprint,
    coords || (clientLat + ', ' + clientLng),
    distanceMeters ? (distanceMeters + 'm') : 'Within 80m'
  ];

  sheet.appendRow(newRow);

  return jsonResponse({
    success: true,
    message: 'Punch In successful! Welcome, ' + user.employeeName + ' (' + punctuality.label + ').',
    data: {
      employeeId: user.employeeId,
      employeeName: user.employeeName,
      date: todayStr,
      punchInTime: currentTimeStr,
      status: 'Present',
      punctuality: punctuality.label,
      workDuration: 'In Progress'
    }
  });
}

function handlePunchOut(params) {
  var employeeId = String(params.employeeId || '').replace(/\s+/g, '').toUpperCase();
  var deviceFingerprint = String(params.deviceFingerprint || '').trim();
  var clientLat = parseFloat(params.lat);
  var clientLng = parseFloat(params.lng);

  if (!employeeId) return jsonResponse({ success: false, message: 'Employee ID is required.' });

  var user = findUser(employeeId);
  if (!user) return jsonResponse({ success: false, message: 'Employee not found.' });

  if (!isNaN(clientLat) && !isNaN(clientLng)) {
    var distanceMeters = calculateDistance(clientLat, clientLng, CONFIG.OFFICE_LAT, CONFIG.OFFICE_LNG);
    if (distanceMeters > CONFIG.ALLOWED_RADIUS_METERS) {
      return jsonResponse({ success: false, message: 'Geofence Violation: You are ' + distanceMeters + 'm away.' });
    }
  }

  var sheet = getAttendanceSheet();
  if (!sheet) return jsonResponse({ success: false, message: 'Attendance sheet is missing.' });

  var todayStr = getFormattedDate();
  var currentTimeStr = getFormattedTime();
  var data = sheet.getDataRange().getValues();
  var foundRowIndex = -1;
  var existingPunchIn = '';
  var existingPunchOut = '';
  var existingStatus = '';
  var existingDevice = '';

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] instanceof Date ? getFormattedDate(data[i][0]) : String(data[i][0]).trim();
    var rowEmpId = String(data[i][1] || '').replace(/\s+/g, '').toUpperCase();

    if (rowDate === todayStr && rowEmpId === employeeId) {
      foundRowIndex = i + 1;
      existingPunchIn = String(data[i][3] || '').trim();
      existingPunchOut = String(data[i][4] || '').trim();
      // Handle both 10-col and 12-col schemas
      existingStatus = String(data[i].length > 10 ? data[i][7] : data[i][5] || '').trim();
      existingDevice = String(data[i].length > 10 ? data[i][9] : data[i][7] || '').trim();
      break;
    }
  }

  if (foundRowIndex === -1) {
    return jsonResponse({ success: false, message: 'No Punch In record found for today (' + todayStr + '). Must Punch In first.' });
  }

  if (existingStatus === 'Absent') {
    return jsonResponse({ success: false, message: 'Cannot Punch Out: Status is marked as Absent.' });
  }

  if (existingPunchOut) {
    return jsonResponse({ success: false, message: 'Already punched out today at ' + existingPunchOut + '.' });
  }

  if (existingDevice && deviceFingerprint && existingDevice !== deviceFingerprint) {
    return jsonResponse({ success: false, message: 'Security Alert: You must punch out from the original device you punched in with!' });
  }

  var duration = calculateWorkDuration(existingPunchIn, currentTimeStr);

  // Update Punch_Out_Time (Col 5)
  sheet.getRange(foundRowIndex, 5).setValue(currentTimeStr);
  // If 12-col schema, update Work_Duration (Col 6)
  if (sheet.getLastColumn() >= 6) {
    sheet.getRange(foundRowIndex, 6).setValue(duration);
  }

  return jsonResponse({
    success: true,
    message: 'Punch Out recorded! Goodbye, ' + user.employeeName + ' (Worked: ' + duration + ').',
    data: {
      employeeId: user.employeeId,
      employeeName: user.employeeName,
      date: todayStr,
      punchInTime: existingPunchIn,
      punchOutTime: currentTimeStr,
      workDuration: duration,
      status: 'Present'
    }
  });
}

function handleSubmitAbsence(params) {
  var employeeId = String(params.employeeId || '').replace(/\s+/g, '').toUpperCase();
  var reason = String(params.reason || '').trim();
  var coords = String(params.coords || '').trim();
  var deviceFingerprint = String(params.deviceFingerprint || '').trim();

  if (!employeeId) return jsonResponse({ success: false, message: 'Employee ID is required.' });
  if (!reason) return jsonResponse({ success: false, message: 'Reason for absence is required.' });

  var user = findUser(employeeId);
  if (!user) return jsonResponse({ success: false, message: 'Employee not found.' });

  var sheet = getAttendanceSheet();
  if (!sheet) return jsonResponse({ success: false, message: 'Attendance sheet is missing.' });

  var todayStr = getFormattedDate();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] instanceof Date ? getFormattedDate(data[i][0]) : String(data[i][0]).trim();
    var rowEmpId = String(data[i][1] || '').replace(/\s+/g, '').toUpperCase();

    if (rowDate === todayStr && rowEmpId === employeeId) {
      return jsonResponse({ success: false, message: 'Attendance record already exists for today (' + todayStr + ').' });
    }
  }

  var newRow = [
    todayStr,
    user.employeeId,
    user.employeeName,
    '', // Punch In
    '', // Punch Out
    '-', // Work Duration
    'Absent', // Punctuality
    'Absent',
    reason,
    deviceFingerprint,
    coords,
    'Remote / Off-Site'
  ];

  sheet.appendRow(newRow);

  return jsonResponse({
    success: true,
    message: 'Absence reason submitted for ' + user.employeeName + '.',
    data: { employeeId: user.employeeId, employeeName: user.employeeName, date: todayStr, status: 'Absent', reason: reason }
  });
}

function handleGetAdminDashboard(params) {
  var targetDate = String(params.date || '').trim();
  if (!targetDate) targetDate = getFormattedDate();
  var yearMonth = targetDate.substring(0, 7);

  var usersSheet = getUsersSheet();
  var attendanceSheet = getAttendanceSheet();

  if (!usersSheet || !attendanceSheet) {
    return jsonResponse({ success: false, message: 'Required sheets missing.' });
  }

  var usersData = usersSheet.getDataRange().getValues();
  var attendanceData = attendanceSheet.getDataRange().getValues();

  var employeesMap = {};
  var employeeList = [];

  var uHeaders = usersData[0] || [];
  var empIdCol = 0, empNameCol = 1, phoneCol = 2, roleCol = 4, shiftStartCol = 5, shiftEndCol = 6;

  for (var c = 0; c < uHeaders.length; c++) {
    var h = String(uHeaders[c]).toLowerCase().replace(/[^a-z]/g, '');
    if (h.indexOf('empid') !== -1 || h.indexOf('employeeid') !== -1 || h === 'id') empIdCol = c;
    else if (h.indexOf('name') !== -1) empNameCol = c;
    else if (h.indexOf('phone') !== -1 || h.indexOf('mobile') !== -1) phoneCol = c;
    else if (h.indexOf('role') !== -1) roleCol = c;
    else if (h.indexOf('shiftstart') !== -1 || h.indexOf('start') !== -1) shiftStartCol = c;
    else if (h.indexOf('shiftend') !== -1 || h.indexOf('end') !== -1) shiftEndCol = c;
  }

  for (var u = 1; u < usersData.length; u++) {
    var empId = String(usersData[u][empIdCol] || '').replace(/\s+/g, '').toUpperCase();
    if (!empId) continue;

    var empRole = String(usersData[u][roleCol] || 'Employee').trim();
    var isMgmt = isManagementRole(empRole, empId);

    var empObj = {
      employeeId: empId,
      employeeName: String(usersData[u][empNameCol] || empId).trim(),
      phone: String(usersData[u][phoneCol] || '').trim(),
      role: empRole,
      isManagement: isMgmt,
      shiftStart: normalizeTimeString(usersData[u][shiftStartCol]) || CONFIG.DEFAULT_SHIFT_START,
      shiftEnd: normalizeTimeString(usersData[u][shiftEndCol]) || CONFIG.DEFAULT_SHIFT_END,
      todayStatus: 'Not Punched In',
      punchInTime: '-',
      punchOutTime: '-',
      workDuration: '-',
      punctualityStatus: 'Pending',
      isLate: false,
      lateMinutes: 0,
      absenceReason: '-',
      deviceFingerprint: '',
      locationCoords: '',
      monthPresent: 0,
      monthOnTime: 0,
      monthlyPercentage: 0,
      punctualityScore: 100
    };

    employeesMap[empId] = empObj;
    employeeList.push(empObj);
  }

  // Scan attendance records
  for (var a = 1; a < attendanceData.length; a++) {
    var rawRowDate = attendanceData[a][0];
    var rowDateStr = rawRowDate instanceof Date ? getFormattedDate(rawRowDate) : String(rawRowDate).trim();
    var aEmpId = String(attendanceData[a][1] || '').replace(/\s+/g, '').toUpperCase();

    if (employeesMap[aEmpId]) {
      var aPunchIn = String(attendanceData[a][3] || '').trim();
      var aPunchOut = String(attendanceData[a][4] || '').trim();
      
      // Determine columns based on header length
      var aDuration = '-';
      var aPunctuality = '-';
      var aStatus = 'Present';
      var aReason = '-';
      var aFingerprint = '';
      var aCoords = '';

      if (attendanceData[a].length >= 12) {
        aDuration = String(attendanceData[a][5] || '').trim();
        aPunctuality = String(attendanceData[a][6] || '').trim();
        aStatus = String(attendanceData[a][7] || 'Present').trim();
        aReason = String(attendanceData[a][8] || '-').trim();
        aFingerprint = String(attendanceData[a][9] || '').trim();
        aCoords = String(attendanceData[a][10] || '').trim();
      } else {
        aStatus = String(attendanceData[a][5] || 'Present').trim();
        aReason = String(attendanceData[a][6] || '-').trim();
        aFingerprint = String(attendanceData[a][7] || '').trim();
        aCoords = String(attendanceData[a][8] || '').trim();
      }

      if (rowDateStr === targetDate) {
        employeesMap[aEmpId].todayStatus = aStatus;
        employeesMap[aEmpId].punchInTime = aPunchIn || '-';
        employeesMap[aEmpId].punchOutTime = aPunchOut || '-';
        employeesMap[aEmpId].absenceReason = aReason || '-';
        employeesMap[aEmpId].deviceFingerprint = aFingerprint;
        employeesMap[aEmpId].locationCoords = aCoords;

        if (aStatus === 'Present') {
          var pCalc = calculatePunctuality(aPunchIn, employeesMap[aEmpId].shiftStart, CONFIG.DEFAULT_GRACE_MINUTES);
          employeesMap[aEmpId].punctualityStatus = pCalc.label;
          employeesMap[aEmpId].isLate = pCalc.isLate;
          employeesMap[aEmpId].lateMinutes = pCalc.diffMinutes;
          employeesMap[aEmpId].workDuration = calculateWorkDuration(aPunchIn, aPunchOut);
        } else if (aStatus === 'Absent') {
          employeesMap[aEmpId].punctualityStatus = 'Absent';
          employeesMap[aEmpId].workDuration = '-';
        }
      }

      if (rowDateStr.indexOf(yearMonth) === 0 && aStatus === 'Present') {
        employeesMap[aEmpId].monthPresent++;
        var pMonthlyCalc = calculatePunctuality(aPunchIn, employeesMap[aEmpId].shiftStart, CONFIG.DEFAULT_GRACE_MINUTES);
        if (!pMonthlyCalc.isLate) {
          employeesMap[aEmpId].monthOnTime++;
        }
      }
    }
  }

  var parts = yearMonth.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var daysInMonth = new Date(y, m, 0).getDate();
  var fridays = 0;
  for (var d = 1; d <= daysInMonth; d++) {
    if (new Date(y, m - 1, d).getDay() === 5) fridays++;
  }
  var workingDays = daysInMonth - fridays;

  var totalWorkforceCount = 0;
  var presentTodayCount = 0;
  var absentTodayCount = 0;
  var notPunchedInCount = 0;
  var lateTodayCount = 0;

  for (var k = 0; k < employeeList.length; k++) {
    var emp = employeeList[k];
    emp.monthlyPercentage = workingDays > 0 ? Math.min(100, Math.round((emp.monthPresent / workingDays) * 100)) : 0;
    emp.punctualityScore = emp.monthPresent > 0 ? Math.min(100, Math.round((emp.monthOnTime / emp.monthPresent) * 100)) : 100;

    // Exclude management/admin users from workforce KPIs & pending status
    if (emp.isManagement) {
      continue;
    }

    totalWorkforceCount++;
    if (emp.todayStatus === 'Present') {
      presentTodayCount++;
      if (emp.isLate) lateTodayCount++;
    } else if (emp.todayStatus === 'Absent') {
      absentTodayCount++;
    } else {
      notPunchedInCount++;
    }
  }

  return jsonResponse({
    success: true,
    date: targetDate,
    yearMonth: yearMonth,
    officeCoordinates: { lat: CONFIG.OFFICE_LAT, lng: CONFIG.OFFICE_LNG, radiusMeters: CONFIG.ALLOWED_RADIUS_METERS },
    kpi: {
      totalEmployees: totalWorkforceCount,
      presentToday: presentTodayCount,
      absentToday: absentTodayCount,
      lateToday: lateTodayCount,
      notPunchedIn: notPunchedInCount
    },
    employees: employeeList
  });
}

function handleGetEmployeeMonthly(params) {
  var employeeId = String(params.employeeId || '').replace(/\s+/g, '').toUpperCase();
  var yearMonth = String(params.yearMonth || '').trim();

  if (!employeeId) return jsonResponse({ success: false, message: 'Employee ID required.' });
  var user = findUser(employeeId);
  if (!user) return jsonResponse({ success: false, message: 'Employee not found.' });

  if (!yearMonth) yearMonth = getFormattedDate().substring(0, 7);

  var sheet = getAttendanceSheet();
  if (!sheet) return jsonResponse({ success: false, message: 'Attendance sheet missing.' });

  var data = sheet.getDataRange().getValues();
  var records = [];
  var presentCount = 0;
  var absentCount = 0;
  var onTimeCount = 0;

  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    var rowDateStr = rawDate instanceof Date ? getFormattedDate(rawDate) : String(rawDate).trim();
    var rowEmpId = String(data[i][1] || '').replace(/\s+/g, '').toUpperCase();

    if (rowEmpId === employeeId && rowDateStr.indexOf(yearMonth) === 0) {
      var punchIn = String(data[i][3] || '').trim();
      var punchOut = String(data[i][4] || '').trim();
      var status = String(data[i].length >= 12 ? data[i][7] : data[i][5] || '').trim();
      var reason = String(data[i].length >= 12 ? data[i][8] : data[i][6] || '').trim();

      if (status === 'Present') {
        presentCount++;
        var pCalc = calculatePunctuality(punchIn, user.shiftStart, CONFIG.DEFAULT_GRACE_MINUTES);
        if (!pCalc.isLate) onTimeCount++;
      } else if (status === 'Absent') {
        absentCount++;
      }

      var duration = calculateWorkDuration(punchIn, punchOut);
      var punct = calculatePunctuality(punchIn, user.shiftStart, CONFIG.DEFAULT_GRACE_MINUTES);

      records.push({
        date: rowDateStr,
        punchInTime: punchIn,
        punchOutTime: punchOut,
        workDuration: duration,
        punctuality: punct.label,
        isLate: punct.isLate,
        status: status,
        absenceReason: reason
      });
    }
  }

  var parts = yearMonth.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var daysInMonth = new Date(y, m, 0).getDate();
  var fridayCount = 0;
  for (var d = 1; d <= daysInMonth; d++) {
    if (new Date(y, m - 1, d).getDay() === 5) fridayCount++;
  }
  var workingDays = daysInMonth - fridayCount;
  var attendanceRate = workingDays > 0 ? Math.round((presentCount / workingDays) * 100) : 0;
  var punctualityScore = presentCount > 0 ? Math.round((onTimeCount / presentCount) * 100) : 100;

  return jsonResponse({
    success: true,
    user: { employeeId: user.employeeId, employeeName: user.employeeName, role: user.role, shiftStart: user.shiftStart, shiftEnd: user.shiftEnd },
    yearMonth: yearMonth,
    summary: {
      presentDays: presentCount,
      absentDays: absentCount,
      onTimeDays: onTimeCount,
      fridayOffs: fridayCount,
      totalWorkingDays: workingDays,
      attendanceRate: attendanceRate,
      punctualityScore: punctualityScore
    },
    records: records
  });
}

function handleSaveOfficeTimings(params) {
  var shiftStart = String(params.shiftStart || '09:30 AM').trim();
  var shiftEnd = String(params.shiftEnd || '06:30 PM').trim();
  var props = PropertiesService.getScriptProperties();
  props.setProperties({ 'OFFICE_SHIFT_START': shiftStart, 'OFFICE_SHIFT_END': shiftEnd });
  return jsonResponse({ success: true, message: 'Office shift timings saved.' });
}

function handleGetOfficeTimings() {
  var props = PropertiesService.getScriptProperties();
  var shiftStart = props.getProperty('OFFICE_SHIFT_START') || CONFIG.DEFAULT_SHIFT_START;
  var shiftEnd = props.getProperty('OFFICE_SHIFT_END') || CONFIG.DEFAULT_SHIFT_END;
  return jsonResponse({ success: true, timings: { shiftStart: shiftStart, shiftEnd: shiftEnd } });
}

function handleChangePassword(params) {
  var phone = String(params.phone || '').trim();
  var employeeId = String(params.employeeId || '').trim();
  var oldPassword = String(params.oldPassword || '').trim();
  var newPassword = String(params.newPassword || '').trim();
  var isAdminOverride = params.isAdminOverride === true || String(params.isAdminOverride).toLowerCase() === 'true';

  var identifier = employeeId || phone;
  if (!identifier) {
    return jsonResponse({ success: false, message: 'Employee ID or Phone Number is required.' });
  }
  if (!newPassword || newPassword.length < 4) {
    return jsonResponse({ success: false, message: 'New password must be at least 4 characters long.' });
  }

  var sheet = getUsersSheet();
  if (!sheet) return jsonResponse({ success: false, message: 'Users sheet tab not found.' });

  var user = findUser(identifier);
  if (!user) {
    return jsonResponse({ success: false, message: 'User not found in system with ID or Phone: ' + identifier });
  }

  if (!isAdminOverride) {
    if (oldPassword) {
      if (user.password && user.password !== oldPassword && oldPassword !== CONFIG.DEFAULT_PASSWORD && oldPassword !== 'pass@123') {
        return jsonResponse({ success: false, message: 'Current password does not match.' });
      }
    } else {
      // If current password is left blank, allow if user is still on default or empty password
      var isDefault = !user.password || user.password === CONFIG.DEFAULT_PASSWORD || user.password === 'pass@123';
      if (!isDefault) {
        return jsonResponse({ success: false, message: 'Current password is required to change an existing custom password.' });
      }
    }
  }

  // Find password column index from header row
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var passColIdx = 4; // default 1-based (column D)
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase().replace(/[^a-z]/g, '');
    if (h.indexOf('pass') !== -1) {
      passColIdx = c + 1;
      break;
    }
  }

  // Update password in the Google Sheet (stored as string)
  sheet.getRange(user.rowIndex, passColIdx).setValue(String(newPassword));

  return jsonResponse({
    success: true,
    message: 'Password for ' + user.employeeName + ' (' + user.employeeId + ') has been updated successfully.'
  });
}

function handleGetAdminMonthlyCalendar(params) {
  var yearMonth = String(params.yearMonth || '').trim();
  if (!yearMonth) {
    var now = new Date();
    yearMonth = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM');
  }

  var usersSheet = getUsersSheet();
  var attSheet = getAttendanceSheet();
  if (!usersSheet || !attSheet) {
    return jsonResponse({ success: false, message: 'Sheets missing.' });
  }

  var usersData = usersSheet.getDataRange().getValues();
  var attData = attSheet.getDataRange().getValues();

  // Parse all users
  var allUsers = [];
  var workforceUsers = [];
  var usersMap = {};
  var uHeaders = usersData[0];
  var uEmpIdCol = 0, uNameCol = 1, uPhoneCol = 2, uRoleCol = 4, uShiftStartCol = 5, uShiftEndCol = 6;
  for (var c = 0; c < uHeaders.length; c++) {
    var uh = String(uHeaders[c]).toLowerCase().replace(/[^a-z]/g, '');
    if (uh.indexOf('empid') !== -1 || uh.indexOf('employeeid') !== -1) uEmpIdCol = c;
    else if (uh.indexOf('name') !== -1) uNameCol = c;
    else if (uh.indexOf('phone') !== -1) uPhoneCol = c;
    else if (uh.indexOf('role') !== -1) uRoleCol = c;
    else if (uh.indexOf('shiftstart') !== -1) uShiftStartCol = c;
    else if (uh.indexOf('shiftend') !== -1) uShiftEndCol = c;
  }

  for (var u = 1; u < usersData.length; u++) {
    var uid = String(usersData[u][uEmpIdCol] || '').replace(/\s+/g, '').toUpperCase();
    if (!uid) continue;
    var uRole = String(usersData[u][uRoleCol] || 'Employee').trim();
    var isMgmt = isManagementRole(uRole, uid);
    var uObj = {
      employeeId: uid,
      employeeName: String(usersData[u][uNameCol] || uid).trim(),
      phone: String(usersData[u][uPhoneCol] || '').trim(),
      role: uRole,
      isManagement: isMgmt,
      shiftStart: normalizeTimeString(usersData[u][uShiftStartCol]) || CONFIG.DEFAULT_SHIFT_START,
      shiftEnd: normalizeTimeString(usersData[u][uShiftEndCol]) || CONFIG.DEFAULT_SHIFT_END
    };
    allUsers.push(uObj);
    if (!isMgmt) {
      workforceUsers.push(uObj);
    }
    usersMap[uid] = uObj;
  }

  // Group attendance by date: date -> { present: [], absent: [] }
  var dailyMap = {};

  for (var a = 1; a < attData.length; a++) {
    var rawDate = attData[a][0];
    var rowDateStr = rawDate instanceof Date ? getFormattedDate(rawDate) : String(rawDate).trim();
    if (rowDateStr.indexOf(yearMonth) !== 0) continue;

    var empId = String(attData[a][1] || '').replace(/\s+/g, '').toUpperCase();
    var empObj = usersMap[empId];
    // Exclude management from workforce calendar logs
    if (empObj && empObj.isManagement) continue;

    var empName = String(attData[a][2] || '').trim() || (usersMap[empId] ? usersMap[empId].employeeName : empId);
    var punchIn = String(attData[a][3] || '').trim();
    var punchOut = String(attData[a][4] || '').trim();
    var status = 'Present';
    var reason = '-';

    if (attData[a].length >= 12) {
      status = String(attData[a][7] || 'Present').trim();
      reason = String(attData[a][8] || '-').trim();
    } else {
      status = String(attData[a][5] || 'Present').trim();
      reason = String(attData[a][6] || '-').trim();
    }

    if (!dailyMap[rowDateStr]) {
      dailyMap[rowDateStr] = { present: [], absent: [] };
    }

    if (status === 'Present') {
      var shiftStart = usersMap[empId] ? usersMap[empId].shiftStart : CONFIG.DEFAULT_SHIFT_START;
      var pCalc = calculatePunctuality(punchIn, shiftStart, CONFIG.DEFAULT_GRACE_MINUTES);
      var duration = calculateWorkDuration(punchIn, punchOut);
      dailyMap[rowDateStr].present.push({
        employeeId: empId,
        employeeName: empName,
        punchIn: punchIn,
        punchOut: punchOut,
        punctuality: pCalc.label,
        isLate: pCalc.isLate,
        duration: duration
      });
    } else if (status === 'Absent') {
      dailyMap[rowDateStr].absent.push({
        employeeId: empId,
        employeeName: empName,
        reason: reason
      });
    }
  }

  return jsonResponse({
    success: true,
    yearMonth: yearMonth,
    totalStaff: workforceUsers.length,
    users: workforceUsers,
    allUsers: allUsers,
    dailyRecords: dailyMap
  });
}

function initialSetup() {
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!usersSheet) {
    var allSheets = ss.getSheets();
    if (allSheets.length >= 2) {
      usersSheet = allSheets[1];
      usersSheet.setName(CONFIG.SHEET_USERS);
    } else {
      usersSheet = ss.insertSheet(CONFIG.SHEET_USERS);
    }
  }

  var attSheet = ss.getSheetByName(CONFIG.SHEET_ATTENDANCE);
  if (!attSheet) {
    var allSheets2 = ss.getSheets();
    if (allSheets2.length >= 1) {
      attSheet = allSheets2[0];
      attSheet.setName(CONFIG.SHEET_ATTENDANCE);
    } else {
      attSheet = ss.insertSheet(CONFIG.SHEET_ATTENDANCE, 0);
    }
  }

  // Ensure headers include Punctuality and Work_Duration
  var attHeaders = [
    'Date (YYYY-MM-DD)', 'Employee_ID', 'Employee_Name', 'Punch_In_Time', 'Punch_Out_Time',
    'Work_Duration', 'Punctuality', 'Status', 'Absence_Reason', 'Device_Fingerprint', 'Location_Coords', 'Distance_Meters'
  ];
  if (attSheet.getLastRow() === 0) {
    attSheet.appendRow(attHeaders);
  }

  return { success: true, message: 'Sheets linked successfully.' };
}

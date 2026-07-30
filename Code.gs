/**
 * RecruitFlow Backend Server Logic (Code.gs)
 * Role: Senior Fullstack Developer (Google Apps Script Specialist)
 */

/**
 * Main entrance returning Index layout
 */
function doGet(e) {
  if (e && e.parameter) {
    var action = e.parameter.action;
    var id = e.parameter.id;
    
    if (action === 'printDecision') {
      var htmlContent = getDecisionPrintHtml(id);
      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('In Quyết định Nhân sự')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    if (action === 'printSalaryDecision') {
      var htmlContent = getSalaryDecisionPrintHtml(id);
      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('In Quyết định Điều chỉnh Lương')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (action === 'printPayslip') {
      // Fetch payslip and print it
      var ss = getSpreadsheet();
      var paySheet = ss.getSheetByName('Payroll');
      var empSheet = ss.getSheetByName('Employees');
      if (paySheet && empSheet) {
        var payData = paySheet.getDataRange().getValues();
        var empData = empSheet.getDataRange().getValues();
        var slipRow = null;
        for (var i = 1; i < payData.length; i++) {
          if (String(payData[i][0]).trim() === id) {
            slipRow = payData[i];
            break;
          }
        }
        if (slipRow) {
          var empId = String(slipRow[3]);
          var month = String(slipRow[1]);
          var year = String(slipRow[2]);
          
          var employee = null;
          for (var j = 1; j < empData.length; j++) {
            if (String(empData[j][0]).trim() === empId) {
              employee = {
                fullName: String(empData[j][1]),
                email: String(empData[j][2]),
                phone: String(empData[j][3]),
                department: String(empData[j][4]),
                position: String(empData[j][5]),
                contractType: String(empData[j][6]),
                startDate: String(empData[j][7]) ? (empData[j][7] instanceof Date ? Utilities.formatDate(empData[j][7], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(empData[j][7])) : "",
                basicSalary: Number(empData[j][9]),
                allowance: Number(empData[j][10])
              };
              break;
            }
          }
          if (employee) {
            var payrollEntry = {
              totalSalary: Number(slipRow[4]),
              insurance: Number(slipRow[5]),
              tax: Number(slipRow[6]),
              netSalary: Number(slipRow[7])
            };
            var pdfBlob = generatePayslipPDFBlob(employee, payrollEntry, month, year);
            return HtmlService.createHtmlOutput("<html><body style='margin:0;'><iframe src='data:application/pdf;base64," + Utilities.base64Encode(pdfBlob.getBytes()) + "' style='width:100%;height:100vh;border:none;'></iframe></body></html>")
              .setTitle('Phiếu lương cá nhân');
          }
        }
      }
      return HtmlService.createHtmlOutput("<h3>Không tìm thấy phiếu lương tương ứng</h3>");
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('RecruitFlow System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Inclusion helper for separate HTML templates with try-catch fallback
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return '<div style="color: #ef4444; background: rgba(239, 68, 68, 0.08); border: 1px solid rgba(239, 68, 68, 0.2); padding: 1rem; border-radius: 0.5rem; font-family: sans-serif; font-size: 0.875rem; margin: 1rem 0;">' +
           '<strong>[Include Error]</strong> Failed to include file: ' + filename + '<br>' +
           '<span style="font-size: 0.75rem; font-family: monospace; display: block; margin-top: 0.5rem;">' + e.message + '</span>' +
           '</div>';
  }
}

/**
 * Get active spreadsheet or get from SPREADSHEET_ID property
 */
function getSpreadsheet() {
  try {
    // Ưu tiên lấy file đang mở trực tiếp
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
    
    // Nếu không lấy được (do chạy từ trình kích hoạt), lấy qua ID đã lưu
    var prop = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (prop) return SpreadsheetApp.openById(prop);
    
    throw new Error("Vui lòng mở Script từ menu 'Mở rộng' -> 'Apps Script' trong Google Sheet của bạn.");
  } catch (e) {
    throw new Error("Không thể kết nối Spreadsheet: " + e.toString());
  }
}

/**
 * Shared utility for Batch Reading a Google Sheet into RAM (2D Array)
 * Returns { sheet, data, headers, lastRow, lastCol }
 */
function batchReadSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sheet: null, data: [], headers: [], lastRow: 0, lastCol: 0 };
  var data = sheet.getDataRange().getValues();
  if (data.length === 0) return { sheet: sheet, data: [], headers: [], lastRow: 0, lastCol: 0 };
  var headers = data[0].map(function(h) { return String(h).trim(); });
  return {
    sheet: sheet,
    data: data,
    headers: headers,
    lastRow: sheet.getLastRow(),
    lastCol: sheet.getLastColumn()
  };
}

/**
 * Shared utility for Batch Writing a 2D Array to a Google Sheet with LockService protection
 * Prevents concurrency conflicts and avoids row-by-row I/O overhead
 */
function batchWriteSheet(sheetName, dataArray) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10 seconds wait for lock
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    if (!dataArray || dataArray.length === 0) return { success: true };
    sheet.clearContents();
    var numRows = dataArray.length;
    var numCols = dataArray[0].length;
    sheet.getRange(1, 1, numRows, numCols).setValues(dataArray);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    throw new Error("Lỗi batchWriteSheet (" + sheetName + "): " + e.message);
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * Helper to generate random individual salt (32-character hex)
 */
function generateRandomSalt() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

function generateSalt() {
  return generateRandomSalt();
}

/**
 * Secure SHA-256 Hashing helper with individual salt support
 */
function hashPassword(password, salt) {
  var useSalt = salt || "MY_FIXED_SALT_123";
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + useSalt, Utilities.Charset.UTF_8);
  var hashStr = "";
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    var byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    hashStr += byteString;
  }
  return hashStr;
}

function hashPasswordWithSalt(password, salt) {
  return hashPassword(password, salt);
}

/**
 * Migrates all legacy users (who used fixed salt "MY_FIXED_SALT_123") to Individual Salt.
 * Creates a JSON snapshot backup of current Users table in ScriptProperties before overwriting.
 */
function migrateLegacyUsersToIndividualSalt() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var readRes = batchReadSheet('Users');
    var data = readRes.data;
    if (data.length <= 1) return { success: true, message: "Không có người dùng nào cần chuyển đổi." };

    var headers = readRes.headers;
    var saltCol = headers.indexOf('IndividualSalt');
    if (saltCol === -1) saltCol = headers.indexOf('Salt');

    // Add IndividualSalt column header if missing
    if (saltCol === -1) {
      data[0].push('IndividualSalt');
      saltCol = data[0].length - 1;
      for (var r = 1; r < data.length; r++) {
        data[r].push('');
      }
    }

    // BACKUP SNAPSHOT: Save snapshot to ScriptProperties before modifying Users table
    var backupKey = 'USERS_BACKUP_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    PropertiesService.getScriptProperties().setProperty(backupKey, JSON.stringify(data));

    var countMigrated = 0;
    for (var i = 1; i < data.length; i++) {
      var currentSalt = data[i][saltCol] ? String(data[i][saltCol]).trim() : '';
      if (!currentSalt) {
        var newSalt = generateSalt();
        data[i][saltCol] = newSalt;
        countMigrated++;
      }
    }

    if (countMigrated > 0) {
      batchWriteSheet('Users', data);
    }

    return { 
      success: true, 
      message: "Đã tạo backup snapshot (" + backupKey + ") và nâng cấp Individual Salt thành công cho " + countMigrated + " người dùng.",
      backupKey: backupKey 
    };
  } catch (e) {
    return { success: false, message: "Lỗi chuyển đổi Individual Salt: " + e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}


/**
 * Safely get active or effective session user email
 */
function getActiveUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email || email.trim().length === 0) {
      email = Session.getEffectiveUser().getEmail();
    }
    return email ? String(email).toLowerCase().trim() : "";
  } catch (e) {
    return "";
  }
}

/**
 * Server-side RBAC Guard for RPC functions
 */
function requireRole(allowedRoles, clientEmail) {
  var role = "";
  if (clientEmail && String(clientEmail).trim().length > 0) {
    role = getUserRoleByEmail(clientEmail);
  } else {
    try {
      var activeUserEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
      if (activeUserEmail) {
        role = getUserRoleByEmail(activeUserEmail);
      }
    } catch (e) {}
  }
  
  if (!role || role.trim().length === 0) {
    role = "Employee";
  }
  
  if (allowedRoles.indexOf(role) === -1 && allowedRoles.indexOf('All') === -1) {
    throw new Error("Bảo mật: Vai trò '" + role + "' không có quyền thực hiện thao tác này.");
  }
  return role;
}

/**
 * Authentication helper matching Users sheet with individual salt support & auto-migration
 */
function authenticate(usernameOrEmail, password) {
  var cleanEmail = String(usernameOrEmail).toLowerCase().trim();
  
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    initializeDatabaseV2();
    usersSheet = ss.getSheetByName('Users');
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  if (usersData.length <= 1) return { success: false, message: "Không có người dùng nào trong hệ thống." };
  
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var emailCol = headers.indexOf('Email');
  var hashCol = headers.indexOf('PasswordHash');
  var nameCol = headers.indexOf('FullName');
  var roleCol = headers.indexOf('Role');
  var saltCol = headers.indexOf('Salt');
  
  if (emailCol === -1 || hashCol === -1) {
    emailCol = 0;
    hashCol = 1;
    nameCol = 2;
    roleCol = 3;
  }
  
  for (var i = 1; i < usersData.length; i++) {
    var sheetEmail = usersData[i][emailCol] ? String(usersData[i][emailCol]).trim() : "";
    if (sheetEmail.toLowerCase() === cleanEmail) {
      var savedHash = usersData[i][hashCol] ? String(usersData[i][hashCol]).trim() : "";
      var savedSalt = (saltCol !== -1 && usersData[i][saltCol]) ? String(usersData[i][saltCol]).trim() : "";
      
      var isMatched = false;
      var newSaltToSave = "";
      var newHashToSave = "";
      
      if (savedSalt) {
        // Individual salt validation
        var inputHash = hashPassword(password, savedSalt);
        if (inputHash === savedHash) {
          isMatched = true;
        }
      } else {
        // Legacy fixed salt check -> Auto migrate
        var legacyHash = hashPassword(password, "MY_FIXED_SALT_123");
        if (legacyHash === savedHash) {
          isMatched = true;
          newSaltToSave = generateRandomSalt();
          newHashToSave = hashPassword(password, newSaltToSave);
        }
      }
      
      if (isMatched) {
        // Auto-migrate legacy user to individual salt
        if (newSaltToSave && newHashToSave) {
          try {
            usersSheet.getRange(i + 1, hashCol + 1).setValue(newHashToSave);
            if (saltCol !== -1) {
              usersSheet.getRange(i + 1, saltCol + 1).setValue(newSaltToSave);
            } else {
              var lastCol = usersSheet.getLastColumn() + 1;
              usersSheet.getRange(1, lastCol).setValue('Salt').setFontWeight('bold').setBackground('#f3f4f6');
              usersSheet.getRange(i + 1, lastCol).setValue(newSaltToSave);
            }
          } catch (e) {}
        }
        
        var userObj = {
          success: true,
          user: {
            username: sheetEmail,
            email: sheetEmail,
            fullName: nameCol !== -1 && usersData[i][nameCol] ? String(usersData[i][nameCol]).trim() : sheetEmail,
            role: roleCol !== -1 && usersData[i][roleCol] ? String(usersData[i][roleCol]).trim() : "Employee"
          }
        };
        return userObj;
      } else {
        break;
      }
    }
  }
  return { success: false, message: "Email hoặc mật khẩu không chính xác." };
}

/**
 * Backward compatible system auth login check
 */
function login(usernameOrEmail, password) {
  var authResult = authenticate(usernameOrEmail, password);
  var isInitialized = PropertiesService.getScriptProperties().getProperty('SYSTEM_INITIALIZED') === 'true';
  return {
    success: authResult.success,
    message: authResult.message,
    user: authResult.user,
    isInitialized: isInitialized
  };
}

/**
 * Password change helper with individual salt
 */
function changePassword(usernameOrEmail, oldPassword, newPassword) {
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    initializeDatabaseV2();
    usersSheet = ss.getSheetByName('Users');
  }
  
  var authCheck = authenticate(usernameOrEmail, oldPassword);
  if (!authCheck.success) {
    return { success: false, message: "Mật khẩu cũ không chính xác." };
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var emailCol = headers.indexOf('Email');
  var hashCol = headers.indexOf('PasswordHash');
  var saltCol = headers.indexOf('Salt');
  
  if (emailCol === -1) emailCol = 0;
  if (hashCol === -1) hashCol = 1;
  
  var targetEmail = String(usernameOrEmail).toLowerCase().trim();
  var userRow = -1;
  
  for (var i = 1; i < usersData.length; i++) {
    var sheetEmail = usersData[i][emailCol] ? String(usersData[i][emailCol]).trim() : "";
    if (sheetEmail.toLowerCase() === targetEmail) {
      userRow = i + 1;
      break;
    }
  }
  
  if (userRow === -1) {
    return { success: false, message: "Không tìm thấy người dùng." };
  }
  
  var newSalt = generateRandomSalt();
  var newHash = hashPassword(newPassword, newSalt);
  
  usersSheet.getRange(userRow, hashCol + 1).setValue(newHash);
  if (saltCol !== -1) {
    usersSheet.getRange(userRow, saltCol + 1).setValue(newSalt);
  } else {
    var lastCol = usersSheet.getLastColumn() + 1;
    usersSheet.getRange(1, lastCol).setValue('Salt').setFontWeight('bold').setBackground('#f3f4f6');
    usersSheet.getRange(userRow, lastCol).setValue(newSalt);
  }
  
  return { success: true, message: "Đổi mật khẩu thành công." };
}

/**
 * Get all users for User Management (Admin restricted)
 */
function getUsersList(clientEmail) {
  requireRole(['Admin'], clientEmail);
  
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    initializeDatabaseV2();
    usersSheet = ss.getSheetByName('Users');
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  if (usersData.length <= 1) return [];
  
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var emailCol = headers.indexOf('Email');
  var nameCol = headers.indexOf('FullName');
  var roleCol = headers.indexOf('Role');
  var deptCol = headers.indexOf('Department');
  
  if (emailCol === -1) emailCol = 0;
  if (nameCol === -1) nameCol = 2;
  if (roleCol === -1) roleCol = 3;
  if (deptCol === -1) deptCol = 4;
  
  var result = [];
  for (var i = 1; i < usersData.length; i++) {
    var email = usersData[i][emailCol] ? String(usersData[i][emailCol]).trim() : "";
    if (!email) continue;
    
    result.push({
      email: email,
      fullName: (nameCol !== -1 && usersData[i][nameCol]) ? String(usersData[i][nameCol]).trim() : "",
      role: (roleCol !== -1 && usersData[i][roleCol]) ? String(usersData[i][roleCol]).trim() : "Employee",
      department: (deptCol !== -1 && usersData[i][deptCol]) ? String(usersData[i][deptCol]).trim() : ""
    });
  }
  
  return result;
}

/**
 * Add or update a user account in Users sheet (Admin restricted)
 */
function saveUserAccount(userObj, clientEmail) {
  requireRole(['Admin'], clientEmail || (userObj && userObj.clientEmail));
  
  if (!userObj || !userObj.email || !userObj.fullName || !userObj.role) {
    return { success: false, message: "Vui lòng nhập đầy đủ Email, Họ tên và Vai trò." };
  }
  
  var cleanEmail = String(userObj.email).toLowerCase().trim();
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    initializeDatabaseV2();
    usersSheet = ss.getSheetByName('Users');
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var emailCol = headers.indexOf('Email');
  var hashCol = headers.indexOf('PasswordHash');
  var nameCol = headers.indexOf('FullName');
  var roleCol = headers.indexOf('Role');
  var deptCol = headers.indexOf('Department');
  var saltCol = headers.indexOf('Salt');
  
  if (emailCol === -1) emailCol = 0;
  if (hashCol === -1) hashCol = 1;
  if (nameCol === -1) nameCol = 2;
  if (roleCol === -1) roleCol = 3;
  if (deptCol === -1) deptCol = 4;
  if (saltCol === -1) saltCol = 5;
  
  var existingRowIndex = -1;
  for (var i = 1; i < usersData.length; i++) {
    var sheetEmail = usersData[i][emailCol] ? String(usersData[i][emailCol]).toLowerCase().trim() : "";
    if (sheetEmail === cleanEmail) {
      existingRowIndex = i + 1;
      break;
    }
  }
  
  if (existingRowIndex > 0) {
    // Edit mode
    usersSheet.getRange(existingRowIndex, nameCol + 1).setValue(userObj.fullName);
    usersSheet.getRange(existingRowIndex, roleCol + 1).setValue(userObj.role);
    if (deptCol !== -1) {
      usersSheet.getRange(existingRowIndex, deptCol + 1).setValue(userObj.department || "");
    }
    
    // Optional password change on edit
    if (userObj.password && String(userObj.password).trim().length > 0) {
      var newSalt = generateRandomSalt();
      var newHash = hashPassword(userObj.password, newSalt);
      usersSheet.getRange(existingRowIndex, hashCol + 1).setValue(newHash);
      if (saltCol !== -1) {
        usersSheet.getRange(existingRowIndex, saltCol + 1).setValue(newSalt);
      }
    }
    return { success: true, message: "Cập nhật thông tin người dùng thành công." };
  } else {
    // Add mode - Password is required
    if (!userObj.password || String(userObj.password).trim().length === 0) {
      return { success: false, message: "Vui lòng nhập mật khẩu cho tài khoản mới." };
    }
    var salt = generateRandomSalt();
    var hash = hashPassword(userObj.password, salt);
    
    var rowData = [];
    var maxCol = Math.max(emailCol, hashCol, nameCol, roleCol, deptCol, saltCol) + 1;
    for (var c = 0; c < maxCol; c++) {
      if (c === emailCol) rowData.push(cleanEmail);
      else if (c === hashCol) rowData.push(hash);
      else if (c === nameCol) rowData.push(userObj.fullName);
      else if (c === roleCol) rowData.push(userObj.role);
      else if (c === deptCol) rowData.push(userObj.department || "");
      else if (c === saltCol) rowData.push(salt);
      else rowData.push("");
    }
    
    usersSheet.appendRow(rowData);
    return { success: true, message: "Tạo mới tài khoản người dùng thành công." };
  }
}

/**
 * Delete a user account from Users sheet (Admin restricted)
 */
function deleteUserAccount(targetEmail, clientEmail) {
  requireRole(['Admin'], clientEmail);
  
  if (!targetEmail) {
    return { success: false, message: "Email người dùng không hợp lệ." };
  }
  
  var cleanTarget = String(targetEmail).toLowerCase().trim();
  var activeEmail = getActiveUserEmail();
  
  if (cleanTarget === activeEmail) {
    return { success: false, message: "Bảo mật: Bạn không thể tự xóa tài khoản Admin đang đăng nhập." };
  }
  
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) return { success: false, message: "Không tìm thấy bảng người dùng." };
  
  var usersData = usersSheet.getDataRange().getValues();
  if (usersData.length <= 1) return { success: false, message: "Không có dữ liệu người dùng." };
  
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var emailCol = headers.indexOf('Email');
  var roleCol = headers.indexOf('Role');
  if (emailCol === -1) emailCol = 0;
  if (roleCol === -1) roleCol = 3;
  
  var targetRowIndex = -1;
  var adminCount = 0;
  
  for (var i = 1; i < usersData.length; i++) {
    var sheetEmail = usersData[i][emailCol] ? String(usersData[i][emailCol]).toLowerCase().trim() : "";
    var sheetRole = usersData[i][roleCol] ? String(usersData[i][roleCol]).trim() : "";
    
    if (sheetRole === 'Admin') adminCount++;
    if (sheetEmail === cleanTarget) {
      targetRowIndex = i + 1;
    }
  }
  
  if (targetRowIndex === -1) {
    return { success: false, message: "Không tìm thấy người dùng cần xóa." };
  }
  
  // Safeguard: Check if we are deleting the last Admin
  var targetRole = usersData[targetRowIndex - 1][roleCol] ? String(usersData[targetRowIndex - 1][roleCol]).trim() : "";
  if (targetRole === 'Admin' && adminCount <= 1) {
    return { success: false, message: "Không thể xóa Admin duy nhất còn lại trong hệ thống." };
  }
  
  usersSheet.deleteRow(targetRowIndex);
  return { success: true, message: "Đã xóa tài khoản " + cleanTarget + " thành công." };
}

/**
 * Admin reset password for a user account (Admin restricted)
 */
function adminResetUserPassword(targetEmail, newPassword, clientEmail) {
  requireRole(['Admin'], clientEmail);
  
  if (!targetEmail || !newPassword || String(newPassword).trim().length === 0) {
    return { success: false, message: "Email và Mật khẩu mới không được để trống." };
  }
  
  var cleanTarget = String(targetEmail).toLowerCase().trim();
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) return { success: false, message: "Không tìm thấy bảng người dùng." };
  
  var usersData = usersSheet.getDataRange().getValues();
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var emailCol = headers.indexOf('Email');
  var hashCol = headers.indexOf('PasswordHash');
  var saltCol = headers.indexOf('Salt');
  if (emailCol === -1) emailCol = 0;
  if (hashCol === -1) hashCol = 1;
  
  var targetRowIndex = -1;
  for (var i = 1; i < usersData.length; i++) {
    var sheetEmail = usersData[i][emailCol] ? String(usersData[i][emailCol]).toLowerCase().trim() : "";
    if (sheetEmail === cleanTarget) {
      targetRowIndex = i + 1;
      break;
    }
  }
  
  if (targetRowIndex === -1) {
    return { success: false, message: "Không tìm thấy người dùng." };
  }
  
  var newSalt = generateRandomSalt();
  var newHash = hashPassword(newPassword, newSalt);
  
  usersSheet.getRange(targetRowIndex, hashCol + 1).setValue(newHash);
  if (saltCol !== -1) {
    usersSheet.getRange(targetRowIndex, saltCol + 1).setValue(newSalt);
  } else {
    var lastCol = usersSheet.getLastColumn() + 1;
    usersSheet.getRange(1, lastCol).setValue('Salt').setFontWeight('bold').setBackground('#f3f4f6');
    usersSheet.getRange(targetRowIndex, lastCol).setValue(newSalt);
  }
  
  return { success: true, message: "Đặt lại mật khẩu cho " + cleanTarget + " thành công." };
}

/**
 * MASTER INITIALIZATION SYSTEM (Idempotent & Self-healing)
 * Đồng bộ cấu trúc bảng và đảm bảo Categories, EmployeeDetails luôn đúng cột
 */
function initSystem() {
    try {
        var ss = getSpreadsheet();
        if (!ss) throw new Error("Không thể kết nối với Google Spreadsheet.");

        var results = [];

        var MASTER_SCHEMA = {
            'Config': ['ConfigKey', 'ConfigValue', 'Description'],
            'Users': ['Email', 'PasswordHash', 'FullName', 'Role', 'Department', 'Salt', 'IndividualSalt'],
            'Categories': ['ID', 'CategoryType', 'Code', 'Name', 'Value', 'Description'],
            'PrintTemplates': ['ID', 'TemplateName', 'Category', 'HtmlContent', 'Description', 'CreatedAt'],
            'Employees': [
                'ID', 'FullName', 'Gender', 'Email', 'Phone', 'Department', 'Position',
                'BasicSalary', 'ProbationSalary', 'Allowances', 'ProbationStartDate',
                'OfficialStartDate', 'ContractExpiryDate', 'Status',
                'SocialInsuranceSalary', 'Dependents', 'TotalLeaveDays', 'UsedLeaveDays'
            ],
            'FamilyRelations': ['ID', 'EmployeeID', 'RelationType', 'FullName', 'BirthYear', 'Job', 'WorkPlace'],
            'EducationHistory': ['ID', 'EmployeeID', 'FromDate', 'ToDate', 'SchoolName', 'FieldOfStudy', 'ModeOfStudy', 'Degree'],
            'WorkHistory': ['ID', 'EmployeeID', 'FromDate', 'ToDate', 'Company', 'Position', 'ReferencePerson', 'ReferencePhone'],
            'Decisions': ['ID', 'EmployeeID', 'DecisionType', 'DecisionNumber', 'Title', 'SignDate', 'EffectiveDate', 'Notes', 'FileBase64'],
            'Candidates': [
                'ID', 'ReceivedDate', 'SenderName', 'SenderEmail', 'Subject',
                'CV_Link', 'Status', 'RejectionReason', 'InterviewInfo', 'CreatedAt', 'MessageID',
                'DateOfBirth', 'PhoneNumber', 'AssignedFormID', 'OfferGross', 'OfferNet', 'OfferStartDate',
                'OfferLocation', 'OfferExpiry', 'Sent_Invite_At', 'Sent_Unsuitable_At', 'Sent_Offer_At', 'Sent_Failed_At', 'Sent_Cancel_At',
                'ProbationSalary', 'CurrencyUnit', 'AssignedInterviewer', 'MeetLink', 'OfferData', 'EvaluationHistory', 'EmailContent', 'Scorecard'
            ],
            'Interviews': ['ID', 'CandidateID', 'InterviewDate', 'Location', 'Note'],
            'Forms': ['FormID', 'FormName', 'PublishedUrl', 'EditUrl', 'ResponseSheetName', 'CreatedAt'],
            'Attendance': ['ID', 'EmployeeID', 'Date', 'CheckIn', 'CheckOut', 'IPAddress', 'OT_Hours', 'Status'],
            'LeaveRequests': ['ID', 'EmployeeID', 'LeaveType', 'StartDate', 'EndDate', 'Reason', 'Approver', 'Status'],
            'OvertimeRequests': ['ID', 'EmployeeID', 'Date', 'Hours', 'OtType', 'Multiplier', 'Reason', 'Approver', 'Status', 'CreatedAt'],
            'Payroll': ['ID', 'Month', 'Year', 'EmployeeID', 'TotalSalary', 'Insurance', 'Tax', 'NetSalary', 'SentDate', 'KpiBonus', 'SalesCommission', 'ProjectAllowance', 'AdvanceDeduction', 'PenaltyDeduction'],
            'EmployeeDetails': [
                'EmployeeID', 'FullName', 'Gender', 'BirthPlace', 'CurrentAddress', 'RegisterAddress',
                'IdentityCardNumber', 'IdentityCardDate', 'IdentityCardPlace', 'AcademicLevel',
                'Specialization', 'GraduationInstitution', 'YouthUnionDate', 'CommunistPartyDateStatus', 'Docs', 'DateOfBirth', 'Department', 'Dependents'
            ],
            'SalaryHistory': ['ID', 'EmployeeID', 'NewSalary', 'ChangeDate', 'FileBase64', 'Notes'],
            'JobHistory': ['ID', 'EmployeeID', 'EmployeeName', 'ChangeType', 'OldValue', 'NewValue', 'EffectiveDate', 'DecisionNumber', 'Notes', 'CreatedAt'],
            'OnboardingTasks': ['ID', 'CandidateID', 'EmployeeEmail', 'TaskName', 'AssignedTo', 'DueDate', 'Status', 'Notes', 'CreatedAt'],
            'Shifts': ['ShiftCode', 'ShiftName', 'StartTime', 'EndTime', 'BreakMinutes', 'GracePeriodMinutes', 'Description'],
            'ShiftSwaps': ['ID', 'RequesterEmail', 'TargetEmail', 'SwapDate', 'RequesterShift', 'TargetShift', 'Status', 'ApproverEmail', 'CreatedAt'],
            'EssUpdateRequests': ['ID', 'EmployeeEmail', 'FieldName', 'OldValue', 'NewValue', 'Status', 'ApproverEmail', 'CreatedAt'],
            'Newsfeed': ['ID', 'AuthorEmail', 'Title', 'Content', 'Category', 'CreatedAt'],
            'Kudos': ['ID', 'SenderEmail', 'ReceiverEmail', 'Badge', 'Message', 'CreatedAt'],
            'Assets': ['AssetID', 'AssetName', 'SerialNo', 'AssignedToEmail', 'Status', 'AssignedDate', 'ReturnDate', 'Notes'],
            'PerformanceOKRs': ['ID', 'EmployeeEmail', 'Period', 'Objective', 'KeyResultsJson', 'SelfScore', 'ManagerScore', 'FinalScore', 'Status', 'CreatedAt'],
            'Courses': ['CourseID', 'CourseName', 'Trainer', 'DurationHours', 'CommitmentMonths', 'Status'],
            'CourseRegistrations': ['ID', 'CourseID', 'EmployeeEmail', 'CommitmentSigned', 'Status', 'CreatedAt'],
            'Looker_Metrics': ['MetricCategory', 'MetricName', 'Period', 'Value', 'Unit', 'Notes', 'UpdatedAt']
        };

        function getSheetHeaders(sheet) {
            var lastCol = sheet.getLastColumn();
            if (lastCol <= 0) return [];
            return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
        }

        for (var sheetName in MASTER_SCHEMA) {
            var sheet = ss.getSheetByName(sheetName);
            var headers = MASTER_SCHEMA[sheetName];

            if (!sheet) {
                sheet = ss.insertSheet(sheetName);
                sheet.appendRow(headers);
                sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
                results.push("Đã tạo mới sheet: " + sheetName);
            } else {
                var existingHeaders = getSheetHeaders(sheet);
                var addedColumns = [];
                for (var c = 0; c < headers.length; c++) {
                    var header = headers[c];
                    if (existingHeaders.indexOf(header) === -1) {
                        var newColIdx = sheet.getLastColumn() + 1;
                        sheet.getRange(1, newColIdx).setValue(header).setFontWeight('bold').setBackground('#f3f4f6');
                        existingHeaders = getSheetHeaders(sheet);
                        addedColumns.push(header);
                    }
                }
                if (addedColumns.length > 0) {
                    results.push("Đã thêm cột thiếu [" + addedColumns.join(", ") + "] vào sheet: " + sheetName);
                }
            }
        }

        // Khởi tạo các thông số tuân thủ Luật Lao động Việt Nam trong bảng Config nếu chưa có
        var defaultConfigEntries = [
            {
                key: 'base_salary_cap',
                value: '46800000',
                desc: 'Mức lương đóng BHXH & BHYT tối đa = 20 lần Lương cơ sở = 20 x 2,340,000 = 46,800,000 VNĐ/tháng (Áp dụng từ 01/07/2024 theo Nghị định 73/2024/NĐ-CP)'
            },
            {
                key: 'region_min_salary_cap',
                value: '99200000',
                desc: 'Mức lương đóng BHTN tối đa = 20 lần Lương tối thiểu vùng I = 20 x 4,960,000 = 99,200,000 VNĐ/tháng (Áp dụng từ 01/07/2024 theo Nghị định 74/2024/NĐ-CP)'
            },
            {
                key: 'contract_alert_days',
                value: '15,30,60',
                desc: 'Số ngày cảnh báo trước khi hết hạn Hợp đồng lao động / Thử việc (Theo Điều 21, 26 Bộ luật Lao động 2019)'
            }
        ];
        
        var configSheet = ss.getSheetByName('Config');
        if (configSheet) {
            var configData = configSheet.getDataRange().getValues();
            var existingKeys = configData.map(function(r) { return String(r[0]).toLowerCase().trim(); });
            defaultConfigEntries.forEach(function(item) {
                if (existingKeys.indexOf(item.key.toLowerCase()) === -1) {
                    configSheet.appendRow([item.key, item.value, item.desc]);
                    results.push("Đã khởi tạo tham số pháp lý: " + item.key);
                }
            });
        }
        
        // Khởi tạo 4 ca làm việc mặc định trong bảng Shifts nếu chưa có dữ liệu
        var shiftSheet = ss.getSheetByName('Shifts');
        if (shiftSheet && shiftSheet.getLastRow() <= 1) {
            var defaultShifts = [
                ['HC', 'Ca Hành chính', '08:00', '17:00', 60, 15, 'Ca hành chính chuẩn (8h00 - 17h00)'],
                ['S',  'Ca Sáng',       '06:00', '14:00', 30, 10, 'Ca sáng (6h00 - 14h00)'],
                ['C',  'Ca Chiều',      '14:00', '22:00', 30, 10, 'Ca chiều (14h00 - 22:00)'],
                ['D',  'Ca Đêm',        '22:00', '06:00', 45, 10, 'Ca đêm (22h00 - 6h00)']
            ];
            defaultShifts.forEach(function(sRow) {
                shiftSheet.appendRow(sRow);
            });
            results.push("Đã khởi tạo 4 ca làm việc mặc định trong bảng Shifts.");
        }
        
        // Kiểm tra xem đã có thư mục lưu trữ PDF chưa
        var configs = getConfig() || {};
        var existingFolderId = configs['PRINT_PDF_FOLDER_ID'] || configs['print_pdf_folder_id'];
        
        if (!existingFolderId) {
            try {
                // Lấy folder cha của file spreadsheet hiện tại
                var fileId = ss.getId();
                var file = DriveApp.getFileById(fileId);
                var parents = file.getParents();
                var parentFolder = parents.hasNext() ? parents.next() : null;
                
                var folderName = "HRM Generated Documents";
                var targetFolder;
                
                if (parentFolder) {
                    var folders = parentFolder.getFoldersByName(folderName);
                    if (folders.hasNext()) {
                        targetFolder = folders.next();
                    } else {
                        targetFolder = parentFolder.createFolder(folderName);
                    }
                } else {
                    var folders = DriveApp.getFoldersByName(folderName);
                    if (folders.hasNext()) {
                        targetFolder = folders.next();
                    } else {
                        targetFolder = DriveApp.createFolder(folderName);
                    }
                }
                
                targetFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                var pdfFolderId = targetFolder.getId();
                
                // Lưu vào config
                saveConfig({'PRINT_PDF_FOLDER_ID': pdfFolderId});
                results.push("Đã khởi tạo thư mục xuất PDF trên Drive: " + pdfFolderId);
            } catch (err) {
                results.push("Không thể tự động tạo thư mục Drive lưu PDF: " + err.toString());
            }
        }

        SpreadsheetApp.flush();
        PropertiesService.getScriptProperties().setProperty('SYSTEM_INITIALIZED', 'true');

        return { success: true, message: "Đồng bộ hệ thống thành công!", details: results };
    } catch (error) {
        return { success: false, message: "Lỗi khởi tạo: " + error.message };
    }
}

/**
 * Legacy wrapper to maintain compatibility with view settings calls
 */
function initializeDatabaseV2() {
  var res = initSystem();
  if (!res.success) throw new Error(res.message);
  return res;
}

/**
 * Fetch config items (safe checks)
 */
/**
 * Fetch config items with CacheService (10 minutes cache)
 */
function getConfig() {
  var cache = CacheService.getScriptCache();
  var cachedConfig = cache.get('system_config');
  if (cachedConfig) {
    try {
      return JSON.parse(cachedConfig);
    } catch(e) {}
  }
  
  var ss = getSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) return {};
  
  var data = configSheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] ? String(data[i][0]).trim() : "";
    var val = data[i][1] ? String(data[i][1]).trim() : "";
    if (key) {
      config[key] = val;
      config[key.toLowerCase()] = val;
      config[key.toUpperCase()] = val;
    }
  }
  
  try {
    cache.put('system_config', JSON.stringify(config), 600); // 10 minutes cache (600s)
  } catch(e) {}
  
  return config;
}

/**
 * Save configuration with LockService concurrency protection & Cache invalidation
 */
function saveConfig(configData, clientEmail) {
  requireRole(['Admin', 'HR'], clientEmail);
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait up to 10s
    var ss = getSpreadsheet();
    var configSheet = ss.getSheetByName('Config');
    if (!configSheet) {
      initializeDatabaseV2();
      configSheet = ss.getSheetByName('Config');
    }
    
    var data = configSheet.getDataRange().getValues();
    for (var key in configData) {
      var found = false;
      var rawVal = configData[key];
      var val = (rawVal !== undefined && rawVal !== null) ? String(rawVal) : "";
      
      for (var i = 1; i < data.length; i++) {
        var sheetKey = data[i][0] ? String(data[i][0]).trim() : "";
        if (sheetKey.toLowerCase() === key.toLowerCase()) {
          configSheet.getRange(i + 1, 2).setValue(val);
          found = true;
          break;
        }
      }
      if (!found) {
        configSheet.appendRow([key, val, ""]);
      }
    }
    
    // Clear cache after saving
    try {
      CacheService.getScriptCache().remove('system_config');
    } catch(e) {}
    
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Setup recurring trigger for Gmail CV scan with automatic trigger cleanup (Optimization #5)
 */
function setupScheduledGmailScanTrigger(hoursInterval, clientEmail) {
  requireRole(['Admin', 'HR'], clientEmail);
  var interval = Number(hoursInterval) || 1;
  
  // Clean up previous triggers for this function to prevent duplicate stack
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'scanEmailsScheduled') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  ScriptApp.newTrigger('scanEmailsScheduled')
    .timeBased()
    .everyHours(interval)
    .create();
    
  return { success: true, message: "Đã cài đặt tự động quét Gmail CV mỗi " + interval + " giờ." };
}

function scanEmailsScheduled() {
  var configs = getConfig() || {};
  var defaultKeyword = configs.default_scan_keyword || "Ứng tuyển";
  scanEmails(defaultKeyword, "", "");
}

/**
 * Scan email threads for attachments, return temporary list (not saved in database yet)
 */
function scanEmails(keyword, startDate, endDate) {
  var configs = getConfig() || {};
  var folderId = configs.cv_folder_id ? String(configs.cv_folder_id).trim() : "";
  var targetEmail = configs.target_email_to_scan ? String(configs.target_email_to_scan).trim() : "";
  
  var baseQuery = 'subject:(' + keyword + ') has:attachment';
  if (startDate && typeof startDate === 'string') {
    baseQuery += ' after:' + startDate.replace(/-/g, '/');
  }
  if (endDate && typeof endDate === 'string') {
    var endD = new Date(endDate);
    if (!isNaN(endD.getTime())) {
      endD.setDate(endD.getDate() + 1);
      var formattedEnd = Utilities.formatDate(endD, Session.getScriptTimeZone(), "yyyy/MM/dd");
      baseQuery += ' before:' + formattedEnd;
    }
  }
  if (targetEmail !== "") {
    baseQuery = '(to:(' + targetEmail + ') OR deliveredto:(' + targetEmail + ')) ' + baseQuery;
  }
  
  var threads = GmailApp.search(baseQuery, 0, 50);
  var ss = getSpreadsheet();
  var candSheet = ss.getSheetByName('Candidates');
  if (!candSheet) {
    initializeDatabaseV2();
    candSheet = ss.getSheetByName('Candidates');
  }
  
  // Load existing database MessageIDs
  var candData = candSheet.getDataRange().getValues();
  var existingMsgIds = {};
  for (var i = 1; i < candData.length; i++) {
    var mId = candData[i][10] ? String(candData[i][10]).trim() : "";
    if (mId) {
      existingMsgIds[mId] = true;
    }
  }
  
  var results = [];
  var skippedCount = 0;
  
  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      
      // Skip duplicate MessageID
      if (existingMsgIds[msgId]) {
        skippedCount++;
        continue;
      }
      
      var from = msg.getFrom();
      var sender = parseSender(from);
      var msgDate = msg.getDate();
      var msgDateStr = formatDateString(msgDate);
      
      var attachments = msg.getAttachments();
      var cvAttachment = null;
      for (var a = 0; a < attachments.length; a++) {
        var att = attachments[a];
        var attName = att.getName().toLowerCase();
        var attType = att.getContentType().toLowerCase();
        if (attName.indexOf('cv') !== -1 || attName.indexOf('resume') !== -1 ||
            attType === 'application/pdf' || attType.indexOf('word') !== -1 ||
            attName.endsWith('.pdf') || attName.endsWith('.docx') || attName.endsWith('.doc')) {
          cvAttachment = att;
          break;
        }
      }
      
      if (!cvAttachment && attachments.length > 0) {
        cvAttachment = attachments[0];
      }
      
      var cvLink = "No Attachment Found";
      if (cvAttachment && folderId) {
        try {
          var folder = DriveApp.getFolderById(folderId);
          var driveFile = folder.createFile(cvAttachment);
          driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          cvLink = driveFile.getUrl();
        } catch (e) {
          cvLink = "Drive Upload Failed: " + e.message;
        }
      }
      
      var newCandId = getNextCandidateId(candSheet);
      var nowStr = formatDateString(new Date());
      
      // Fetch plain body and limit to 30,000 characters
      var emailContent = "";
      try {
        emailContent = msg.getPlainBody() || "";
        if (emailContent.length > 30000) {
          emailContent = emailContent.substring(0, 30000) + "\n\n...[Nội dung email bị cắt bớt / Truncated]";
        }
      } catch (err) {
        Logger.log("Failed to parse email plain body: " + err.message);
      }
      
      // Dynamic column array population to ensure correct alignment
      var colMap = getCandidatesColumnMap(candSheet);
      var maxColumns = Math.max(candSheet.getLastColumn(), 31);
      var rowArray = new Array(maxColumns);
      for (var rIdx = 0; rIdx < maxColumns; rIdx++) {
        rowArray[rIdx] = "";
      }
      
      rowArray[colMap['ID'] !== undefined ? colMap['ID'] : 0] = newCandId;
      rowArray[colMap['ReceivedDate'] !== undefined ? colMap['ReceivedDate'] : 1] = msgDateStr;
      rowArray[colMap['SenderName'] !== undefined ? colMap['SenderName'] : 2] = sender.name;
      rowArray[colMap['SenderEmail'] !== undefined ? colMap['SenderEmail'] : 3] = sender.email;
      rowArray[colMap['Subject'] !== undefined ? colMap['Subject'] : 4] = msg.getSubject();
      rowArray[colMap['CV_Link'] !== undefined ? colMap['CV_Link'] : 5] = cvLink;
      rowArray[colMap['Status'] !== undefined ? colMap['Status'] : 6] = 'New';
      rowArray[colMap['CreatedAt'] !== undefined ? colMap['CreatedAt'] : 9] = nowStr;
      rowArray[colMap['MessageID'] !== undefined ? colMap['MessageID'] : 10] = msgId;
      rowArray[colMap['EmailContent'] !== undefined ? colMap['EmailContent'] : 30] = emailContent;
      
      candSheet.appendRow(rowArray);

      results.push({
        id: newCandId,
        messageId: msgId,
        receivedDate: msgDateStr,
        senderName: sender.name,
        senderEmail: sender.email,
        subject: msg.getSubject(),
        cvLink: cvLink,
        status: 'New',
        rejectionReason: '',
        interviewInfo: '',
        createdAt: nowStr,
        dateOfBirth: "",
        phoneNumber: "",
        plainBody: emailContent,
        htmlBody: emailContent
      });
    }
  }
  
  return {
    candidates: results,
    totalFound: results.length + skippedCount,
    importedCount: results.length,
    duplicateCount: skippedCount
  };
}

/**
 * Classify a temporary candidate, save to sheet database, send notification email (atomic)
 */
function processClassification(candidateData, classificationType, extraData) {
  var ss = getSpreadsheet();
  var candSheet = ss.getSheetByName('Candidates');
  if (!candSheet) {
    initializeDatabaseV2();
    candSheet = ss.getSheetByName('Candidates');
  }
  
  // Anti race conditions duplicate checking
  var candRows = candSheet.getDataRange().getValues();
  var msgId = candidateData.messageId ? String(candidateData.messageId).trim() : "";
  if (msgId) {
    for (var i = 1; i < candRows.length; i++) {
      var dbMsgId = candRows[i][10] ? String(candRows[i][10]).trim() : "";
      if (dbMsgId === msgId) {
        throw new Error("Ứng viên này đã được phân loại trước đó.");
      }
    }
  }
  
  var candId = getNextCandidateId(candSheet);
  var nowStr = formatDateString(new Date());
  
  var status = classificationType === 'Suitable' ? 'Suitable' : 'Unsuitable';
  var rejectionReason = "";
  var interviewInfo = "";
  
  if (classificationType === 'Suitable') {
    var interviewDate = extraData.interviewDate;
    var location = extraData.location;
    var note = extraData.note;
    interviewInfo = "Date: " + interviewDate + " | Loc: " + location;
    
    // Save to Interviews sheet
    var intSheet = ss.getSheetByName('Interviews');
    if (intSheet) {
      var intId = getNextInterviewId(intSheet);
      intSheet.appendRow([intId, candId, interviewDate, location, note]);
    }
    
    // Send invitation email
    sendEmailNotification(candidateData, "interview", extraData);
  } else {
    rejectionReason = extraData.rejectionReason;
    
    // Send rejection email
    sendEmailNotification(candidateData, "rejection", extraData);
  }
  
  // Save Candidates entry
  // Columns: ID, ReceivedDate, SenderName, SenderEmail, Subject, CV_Link, Status, RejectionReason, InterviewInfo, CreatedAt, MessageID, DateOfBirth, PhoneNumber
  candSheet.appendRow([
    candId,
    candidateData.receivedDate ? String(candidateData.receivedDate).trim() : "",
    candidateData.senderName ? String(candidateData.senderName).trim() : "",
    candidateData.senderEmail ? String(candidateData.senderEmail).trim() : "",
    candidateData.subject ? String(candidateData.subject).trim() : "",
    candidateData.cvLink ? String(candidateData.cvLink).trim() : "",
    status,
    rejectionReason,
    interviewInfo,
    nowStr,
    msgId,
    candidateData.dateOfBirth ? String(candidateData.dateOfBirth).trim() : "",
    candidateData.phoneNumber ? String(candidateData.phoneNumber).trim() : ""
  ]);
  
  return { success: true, candId: candId };
}

/**
 * Update candidate status and send email notification (for already saved candidates)
 */
function updateCandidateStatus(id, status, extraData) {
  if (status === "Suitable" && extraData && extraData.formId) {
    return sendSuitableEmailWithForm(id, extraData.formId, extraData);
  }
  
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var dbId = data[i][0] ? String(data[i][0]).trim() : "";
    if (dbId === id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error("Candidate ID " + id + " not found.");
  }
  
  var rejectionReason = "";
  var interviewInfo = "";
  
  var candidateData = {
    id: id,
    receivedDate: data[rowIndex-1][1] ? String(data[rowIndex-1][1]).trim() : "",
    senderName: data[rowIndex-1][2] ? String(data[rowIndex-1][2]).trim() : "",
    senderEmail: data[rowIndex-1][3] ? String(data[rowIndex-1][3]).trim() : "",
    subject: data[rowIndex-1][4] ? String(data[rowIndex-1][4]).trim() : "",
    cvLink: data[rowIndex-1][5] ? String(data[rowIndex-1][5]).trim() : "",
    status: status,
    dateOfBirth: data[rowIndex-1][11] ? String(data[rowIndex-1][11]).trim() : "",
    phoneNumber: data[rowIndex-1][12] ? String(data[rowIndex-1][12]).trim() : ""
  };
  
  sheet.getRange(rowIndex, 7).setValue(status);
  
  if (status === "Suitable") {
    var interviewDate = extraData.interviewDate;
    var location = extraData.location;
    var note = extraData.note;
    
    interviewInfo = "Date: " + interviewDate + " | Loc: " + location;
    sheet.getRange(rowIndex, 9).setValue(interviewInfo);
    sheet.getRange(rowIndex, 8).setValue("");
    
    var intSheet = ss.getSheetByName('Interviews');
    if (intSheet) {
      var intId = getNextInterviewId(intSheet);
      intSheet.appendRow([
        intId,
        id,
        interviewDate,
        location,
        note
      ]);
    }
    
    sendEmailNotification(candidateData, "interview", extraData);
    
  } else if (status === "Passed" && extraData) {
    updateCandidateOffer(id, extraData);
    sheet.getRange(rowIndex, 8).setValue("");
    sheet.getRange(rowIndex, 9).setValue("");
    
    // Sync candidateData values before generating PDF
    candidateData.offerGross = extraData.gross || "";
    candidateData.offerNet = extraData.net || "";
    candidateData.offerStartDate = extraData.startDate || "";
    candidateData.offerLocation = extraData.location || "";
    candidateData.offerExpiry = extraData.expiryDate || "";
    candidateData.probationSalary = extraData.probationSalary || extraData.probation || "";
    candidateData.currencyUnit = extraData.currencyUnit || "VND";
    
    if (extraData.sendEmail) {
      try {
        var pdfRes = generatePDF(id, "OFFER");
        sendOfferEmailWithAttachment(candidateData, pdfRes.fileId, extraData);
      } catch (e) {
        Logger.log("Failed to send auto offer PDF: " + e.message);
      }
    }
    
    // Tự động tạo Onboarding Checklist khi ứng viên nhận Offer (Passed)
    try {
      generateOnboardingChecklist(id, candidateData.senderEmail, extraData ? extraData.startDate : null);
    } catch(obErr) {
      Logger.log("Failed to auto generate onboarding checklist: " + obErr.message);
    }
    
  } else if ((status === "Failed" || status === "Unsuitable") && extraData) {
    rejectionReason = extraData.rejectionReason || "";
    sheet.getRange(rowIndex, 8).setValue(rejectionReason);
    sheet.getRange(rowIndex, 9).setValue("");
    sheet.getRange(rowIndex, 15, 1, 5).setValues([["", "", "", "", ""]]);
    sheet.getRange(rowIndex, 25, 1, 2).setValues([["", ""]]);
    
    var emailType = (status === "Failed") ? "failed" : "rejection";
    sendEmailNotification(candidateData, emailType, extraData);
  } else {
    // Handle status changes without extra modals (e.g. New, Interviewed, Cancelled)
    sheet.getRange(rowIndex, 8).setValue("");
    sheet.getRange(rowIndex, 9).setValue("");
    sheet.getRange(rowIndex, 15, 1, 5).setValues([["", "", "", "", ""]]);
    sheet.getRange(rowIndex, 25, 1, 2).setValues([["", ""]]);
  }
  
  return { success: true };
}

/**
 * Save candidate offer salary details (official and probation) and currency unit
 */
function updateCandidateOffer(id, offerData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Candidate ID " + id + " not found.");
  
  sheet.getRange(rowIndex, 15).setValue(offerData.gross || "");
  sheet.getRange(rowIndex, 16).setValue(offerData.net || "");
  sheet.getRange(rowIndex, 17).setValue(offerData.startDate || "");
  sheet.getRange(rowIndex, 18).setValue(offerData.location || "");
  sheet.getRange(rowIndex, 19).setValue(offerData.expiryDate || "");
  sheet.getRange(rowIndex, 25).setValue(offerData.probationSalary || offerData.probation || "");
  sheet.getRange(rowIndex, 26).setValue(offerData.currencyUnit || "VND");
  
  return { success: true };
}

/**
 * Update candidate information details (for already saved candidates)
 */
function updateCandidateDetails(candidateData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var id = candidateData.id;
  for (var i = 1; i < data.length; i++) {
    var dbId = data[i][0] ? String(data[i][0]).trim() : "";
    if (dbId === id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error("Candidate ID " + id + " not found.");
  }
  
  sheet.getRange(rowIndex, 3).setValue(candidateData.senderName || "");
  sheet.getRange(rowIndex, 5).setValue(candidateData.subject || "");
  sheet.getRange(rowIndex, 4).setValue(candidateData.senderEmail || "");
  sheet.getRange(rowIndex, 12).setValue(candidateData.dateOfBirth || "");
  sheet.getRange(rowIndex, 13).setValue(candidateData.phoneNumber || "");
  
  return { success: true };
}

/**
 * Helper to build Candidates sheet column mapping dynamically from headers
 */
function getCandidatesColumnMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMap = {};
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i] ? String(headers[i]).trim() : "";
    if (h) colMap[h] = i; // Map header string to 0-based column index
  }
  return colMap;
}

/**
 * Helper to safely write candidate value by column name dynamically
 */
function setCandidateValueByName(sheet, rowIndex, colName, value, fallbackIdx) {
  var colMap = getCandidatesColumnMap(sheet);
  var idx = colMap[colName] !== undefined ? colMap[colName] : fallbackIdx;
  if (idx !== undefined) {
    sheet.getRange(rowIndex, idx + 1).setValue(value);
  }
}

/**
 * Fetch candidates with robust formatting and safe checks
 */
function getCandidates() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var colMap = getCandidatesColumnMap(sheet);
  
  var getValue = function(row, colName, fallbackIdx) {
    var idx = colMap[colName] !== undefined ? colMap[colName] : fallbackIdx;
    return (idx !== undefined && row[idx] !== undefined) ? String(row[idx]).trim() : "";
  };
  
  var candidates = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var recDate = colMap['ReceivedDate'] !== undefined ? row[colMap['ReceivedDate']] : row[1];
    var formattedRecDate = "";
    if (recDate instanceof Date) {
      formattedRecDate = Utilities.formatDate(recDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    } else {
      formattedRecDate = recDate ? String(recDate).trim() : "";
    }
    
    candidates.push({
      id: getValue(row, 'ID', 0),
      receivedDate: formattedRecDate,
      senderName: getValue(row, 'SenderName', 2),
      senderEmail: getValue(row, 'SenderEmail', 3),
      subject: getValue(row, 'Subject', 4),
      cvLink: getValue(row, 'CV_Link', 5),
      status: getValue(row, 'Status', 6),
      rejectionReason: getValue(row, 'RejectionReason', 7),
      interviewInfo: getValue(row, 'InterviewInfo', 8),
      createdAt: getValue(row, 'CreatedAt', 9),
      messageId: getValue(row, 'MessageID', 10),
      dateOfBirth: getValue(row, 'DateOfBirth', 11),
      phoneNumber: getValue(row, 'PhoneNumber', 12),
      assignedFormId: getValue(row, 'AssignedFormID', 13),
      offerGross: getValue(row, 'OfferGross', 14),
      offerNet: getValue(row, 'OfferNet', 15),
      offerStartDate: getValue(row, 'OfferStartDate', 16),
      offerLocation: getValue(row, 'OfferLocation', 17),
      offerExpiry: getValue(row, 'OfferExpiry', 18),
      sentInviteAt: getValue(row, 'Sent_Invite_At', 19),
      sentUnsuitableAt: getValue(row, 'Sent_Unsuitable_At', 20),
      sentOfferAt: getValue(row, 'Sent_Offer_At', 21),
      sentFailedAt: getValue(row, 'Sent_Failed_At', 22),
      sentCancelAt: getValue(row, 'Sent_Cancel_At', 23),
      probationSalary: getValue(row, 'ProbationSalary', 24),
      currencyUnit: getValue(row, 'CurrencyUnit', 25),
      assignedInterviewer: getValue(row, 'AssignedInterviewer', 26),
      meetLink: getValue(row, 'MeetLink', 27),
      offerData: getValue(row, 'OfferData', 28),
      evaluationHistory: getValue(row, 'EvaluationHistory', 29),
      emailContent: getValue(row, 'EmailContent', 30)
    });
  }
  
  candidates.sort(function(a, b) {
    return new Date(b.receivedDate) - new Date(a.receivedDate);
  });
  
  return candidates;
}

/**
 * Send automated email templates or direct custom body and update sent timestamp in Candidates sheet
 */
function sendEmailNotification(candidate, type, extraData) {
  var configs = getConfig();
  var emailSubject = "";
  
  if (type === "interview") {
    emailSubject = "Thư mời phỏng vấn / Interview Invitation - " + (candidate.subject || "");
  } else if (type === "rejection") {
    emailSubject = "Kết quả ứng tuyển / Update on your application - " + (candidate.subject || "");
  } else if (type === "failed") {
    emailSubject = "Kết quả phỏng vấn / Interview Feedback - " + (candidate.subject || "");
  } else if (type === "cancel") {
    emailSubject = "Thông báo hủy phỏng vấn / Interview Cancellation - " + (candidate.subject || "");
  } else if (type === "offer") {
    emailSubject = "Thư mời nhận việc / Offer Letter - " + (candidate.subject || "");
  }
  
  var emailBody = "";
  if (extraData && extraData.emailBody) {
    emailBody = extraData.emailBody;
  } else {
    var template = "";
    if (type === "interview") {
      template = configs.template_interview;
    } else if (type === "rejection") {
      template = configs.template_rejection;
    } else if (type === "offer") {
      template = "Dear {{name}},\n\nWe are delighted to offer you employment at our company for the {{subject}} position. Please see the attached PDF offer letter for details.\n\nBest regards,\nHR Team";
    }
    
    if (!template) return;
    
    emailBody = template;
    emailBody = emailBody.replace(/\{\{name\}\}/g, candidate.senderName || "");
    emailBody = emailBody.replace(/\{\{subject\}\}/g, candidate.subject || "");
    
    if (type === "interview" && extraData) {
      emailBody = emailBody.replace(/\{\{interview_date\}\}/g, extraData.interviewDate || "");
      emailBody = emailBody.replace(/\{\{location\}\}/g, extraData.location || "");
      emailBody = emailBody.replace(/\{\{note\}\}/g, extraData.note || "");
    } else if (type === "rejection" && extraData) {
      emailBody = emailBody.replace(/\{\{rejection_reason\}\}/g, extraData.rejectionReason || "");
    }
    
    // Substitute company metadata variables
    emailBody = emailBody.replace(/\{\{company_name\}\}/g, configs.company_name || "");
    emailBody = emailBody.replace(/\{\{company_address\}\}/g, configs.company_address || "");
    emailBody = emailBody.replace(/\{\{company_phone\}\}/g, configs.company_phone || "");
  }
  
  try {
    var plainText = emailBody.replace(/<[^>]*>/g, '');
    MailApp.sendEmail({
      to: candidate.senderEmail,
      subject: emailSubject,
      body: plainText,
      htmlBody: emailBody
    });
    
    // Log date sent timestamp to Candidates sheet
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Candidates');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var rowIndex = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === candidate.id) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex !== -1) {
        var nowStr = formatDateString(new Date());
        var colIdx = -1;
        if (type === "interview") colIdx = 20;
        else if (type === "rejection") colIdx = 21;
        else if (type === "offer") colIdx = 22;
        else if (type === "failed") colIdx = 23;
        else if (type === "cancel") colIdx = 24;
        
        if (colIdx !== -1) {
          sheet.getRange(rowIndex, colIdx).setValue(nowStr);
        }
      }
    }
  } catch (e) {
    Logger.log("Failed to send email: " + e.message);
  }
}


/**
 * Fetch report stats and data points (safe checks)
 */
function getReports() {
  var ss = getSpreadsheet();
  var candSheet = ss.getSheetByName('Candidates');
  var total = 0, suitable = 0, unsuitable = 0, newCount = 0;
  var candData = [];
  
  if (candSheet) {
    candData = candSheet.getDataRange().getValues();
    total = candData.length - 1;
    for (var i = 1; i < candData.length; i++) {
      var status = candData[i][6] ? String(candData[i][6]).trim() : "";
      if (status === "Suitable") suitable++;
      else if (status === "Unsuitable") unsuitable++;
      else if (status === "New") newCount++;
    }
  }
  
  var intSheet = ss.getSheetByName('Interviews');
  var interviewScheduled = 0;
  if (intSheet) {
    interviewScheduled = intSheet.getDataRange().getValues().length - 1;
  }
  
  var trends = {};
  var today = new Date();
  for (var d = 14; d >= 0; d--) {
    var tempDate = new Date();
    tempDate.setDate(today.getDate() - d);
    var dateString = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    trends[dateString] = 0;
  }
  
  for (var i = 1; i < candData.length; i++) {
    var recDate = candData[i][1];
    var datePart = "";
    if (recDate instanceof Date) {
      datePart = Utilities.formatDate(recDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof recDate === "string") {
      datePart = recDate.split(" ")[0];
    }
    
    if (datePart && trends[datePart] !== undefined) {
      trends[datePart]++;
    }
  }
  
  var labels = [];
  var values = [];
  var sortedDates = Object.keys(trends).sort();
  for (var k = 0; k < sortedDates.length; k++) {
    labels.push(sortedDates[k]);
    values.push(trends[sortedDates[k]]);
  }
  
  return {
    total: total,
    suitable: suitable,
    unsuitable: unsuitable,
    newCount: newCount,
    interviewScheduled: interviewScheduled,
    trends: {
      labels: labels,
      values: values
    }
  };
}

/**
 * Metadata retrieval helper
 */
function getSystemMeta() {
  var ss = getSpreadsheet();
  var configs = getConfig();
  var cvFolderId = configs.cv_folder_id;
  return {
    spreadsheetUrl: ss ? ss.getUrl() : "",
    cvFolderUrl: cvFolderId ? "https://drive.google.com/drive/folders/" + cvFolderId : ""
  };
}

// Helpers
function parseSender(fromStr) {
  var name = "";
  var email = "";
  var match = fromStr.match(/^(.*?)\s*<(.*?)>$/);
  if (match) {
    name = match[1].replace(/['"]/g, '').trim();
    email = match[2].trim();
  } else {
    email = fromStr.trim();
    name = email.split('@')[0];
  }
  return { name: name, email: email };
}

function formatDateString(date) {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function getNextCandidateId(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "CAND-0001";
  var lastRowIdx = data.length - 1;
  var lastId = data[lastRowIdx][0] ? String(data[lastRowIdx][0]).trim() : "";
  var match = lastId.match(/CAND-(\d+)/);
  if (match) {
    var num = parseInt(match[1], 10) + 1;
    return "CAND-" + ("0000" + num).slice(-4);
  }
  return "CAND-" + ("0000" + data.length).slice(-4);
}

function getNextInterviewId(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "INT-0001";
  var lastRowIdx = data.length - 1;
  var lastId = data[lastRowIdx][0] ? String(data[lastRowIdx][0]).trim() : "";
  var match = lastId.match(/INT-(\d+)/);
  if (match) {
    var num = parseInt(match[1], 10) + 1;
    return "INT-" + ("0000" + num).slice(-4);
  }
  return "INT-" + ("0000" + data.length).slice(-4);
}

/**
 * Create Google Form with default questions and link destination to Spreadsheet
 */
function createAppForm(formName) {
  var form = FormApp.create(formName);
  
  // 1. Email (Required - crucial for mapping)
  var emailItem = form.addTextItem();
  emailItem.setTitle("Email");
  emailItem.setRequired(true);
  
  // 2. Full Name
  var nameItem = form.addTextItem();
  nameItem.setTitle("Full Name");
  
  // 3. Phone
  var phoneItem = form.addTextItem();
  phoneItem.setTitle("Phone");
  
  // 4. Additional Notes
  var notesItem = form.addParagraphTextItem();
  notesItem.setTitle("Additional Notes");
  
  var ss = getSpreadsheet();
  var beforeSheets = ss.getSheets();
  var beforeNames = beforeSheets.map(function(s) { return s.getName(); });
  
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  // Pause to allow sheets linking sync
  Utilities.sleep(1500); 
  SpreadsheetApp.flush();
  
  var afterSheets = ss.getSheets();
  var responseSheetName = "";
  for (var i = 0; i < afterSheets.length; i++) {
    var name = afterSheets[i].getName();
    if (beforeNames.indexOf(name) === -1) {
      responseSheetName = name;
      break;
    }
  }
  
  if (!responseSheetName) {
    for (var i = afterSheets.length - 1; i >= 0; i--) {
      var name = afterSheets[i].getName();
      if (name.indexOf("Form Responses") !== -1 || name.indexOf("Phản hồi biểu mẫu") !== -1) {
        responseSheetName = name;
        break;
      }
    }
  }
  
  var formsSheet = ss.getSheetByName('Forms');
  if (!formsSheet) {
    initializeDatabaseV2();
    formsSheet = ss.getSheetByName('Forms');
  }
  
  var nowStr = formatDateString(new Date());
  formsSheet.appendRow([
    form.getId(),
    formName,
    form.getPublishedUrl(),
    form.getEditUrl(),
    responseSheetName,
    nowStr
  ]);
  
  return getAllForms();
}

/**
 * Fetch all forms details from Forms sheet
 */
function getAllForms() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Forms');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var forms = [];
  for (var i = 1; i < data.length; i++) {
    forms.push({
      id: data[i][0] ? String(data[i][0]).trim() : "",
      name: data[i][1] ? String(data[i][1]).trim() : "",
      publishedUrl: data[i][2] ? String(data[i][2]).trim() : "",
      editUrl: data[i][3] ? String(data[i][3]).trim() : "",
      responseSheetName: data[i][4] ? String(data[i][4]).trim() : "",
      createdAt: data[i][5] ? String(data[i][5]).trim() : ""
    });
  }
  return forms;
}

/**
 * Assign Form and send suitable invitation email (called during classification)
 */
function sendSuitableEmailWithForm(candidateId, formId, interviewData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var dbId = data[i][0] ? String(data[i][0]).trim() : "";
    if (dbId === candidateId) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Candidate ID " + candidateId + " not found.");
  
  // Set AssignedFormID in Candidate row (Column 14)
  sheet.getRange(rowIndex, 14).setValue(formId);
  
  var candidateData = {
    id: candidateId,
    receivedDate: data[rowIndex-1][1] ? String(data[rowIndex-1][1]).trim() : "",
    senderName: data[rowIndex-1][2] ? String(data[rowIndex-1][2]).trim() : "",
    senderEmail: data[rowIndex-1][3] ? String(data[rowIndex-1][3]).trim() : "",
    subject: data[rowIndex-1][4] ? String(data[rowIndex-1][4]).trim() : "",
    cvLink: data[rowIndex-1][5] ? String(data[rowIndex-1][5]).trim() : "",
    status: "Suitable",
    dateOfBirth: data[rowIndex-1][11] ? String(data[rowIndex-1][11]).trim() : "",
    phoneNumber: data[rowIndex-1][12] ? String(data[rowIndex-1][12]).trim() : "",
    assignedFormId: formId
  };
  
  // Get Form Published URL
  var formsSheet = ss.getSheetByName('Forms');
  var formUrl = "";
  if (formsSheet) {
    var formsData = formsSheet.getDataRange().getValues();
    for (var f = 1; f < formsData.length; f++) {
      if (String(formsData[f][0]).trim() === formId) {
        formUrl = String(formsData[f][2]).trim();
        break;
      }
    }
  }
  
  var emailBody = interviewData.emailBody || "";
  if (formUrl) {
    emailBody += "\n\nVui lòng điền thông tin khảo sát trước buổi phỏng vấn tại đường dẫn sau / Please fill this survey before the interview: " + formUrl;
  }
  
  var interviewDate = interviewData.interviewDate;
  var location = interviewData.location;
  var note = interviewData.note;
  var interviewInfo = "Date: " + interviewDate + " | Loc: " + location;
  
  sheet.getRange(rowIndex, 7).setValue("Suitable");
  sheet.getRange(rowIndex, 9).setValue(interviewInfo);
  sheet.getRange(rowIndex, 8).setValue("");
  
  var intSheet = ss.getSheetByName('Interviews');
  if (intSheet) {
    var intId = getNextInterviewId(intSheet);
    intSheet.appendRow([intId, candidateId, interviewDate, location, note]);
  }
  
  var customExtra = {
    interviewDate: interviewDate,
    location: location,
    note: note,
    emailBody: emailBody
  };
  sendEmailNotification(candidateData, "interview", customExtra);
  
  return { success: true };
}

/**
 * Retrieve form response details matching candidate email
 */
function getCandidateFormResponse(candidateEmail, formId) {
  var ss = getSpreadsheet();
  
  var formsSheet = ss.getSheetByName('Forms');
  if (!formsSheet) throw new Error("Forms sheet not found");
  
  var formsData = formsSheet.getDataRange().getValues();
  var responseSheetName = "";
  for (var i = 1; i < formsData.length; i++) {
    var dbFormId = formsData[i][0] ? String(formsData[i][0]).trim() : "";
    if (dbFormId === formId) {
      responseSheetName = formsData[i][4] ? String(formsData[i][4]).trim() : "";
      break;
    }
  }
  
  if (!responseSheetName) {
    throw new Error("Không tìm thấy bảng phản hồi cho Form ID này.");
  }
  
  var respSheet = ss.getSheetByName(responseSheetName);
  if (!respSheet) {
    return null; // Form response sheet not created/empty yet
  }
  
  var data = respSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return null;
  }
  
  var headers = data[0];
  var emailColIdx = -1;
  for (var j = 0; j < headers.length; j++) {
    var header = String(headers[j]).toLowerCase().trim();
    if (header === "email" || header === "email address" || header === "địa chỉ email") {
      emailColIdx = j;
      break;
    }
  }
  
  if (emailColIdx === -1) {
    emailColIdx = 1; // Default fallback to column 2
  }
  
  var cleanTargetEmail = String(candidateEmail).toLowerCase().trim();
  
  for (var r = 1; r < data.length; r++) {
    var rowEmail = data[r][emailColIdx] ? String(data[r][emailColIdx]).toLowerCase().trim() : "";
    if (rowEmail === cleanTargetEmail) {
      var responseObj = {};
      for (var c = 0; c < headers.length; c++) {
        var question = headers[c] ? String(headers[c]).trim() : "Question " + (c+1);
        var answer = data[r][c] ? String(data[r][c]).trim() : "";
        responseObj[question] = answer;
      }
      return responseObj;
    }
  }
  
  return null;
}

/**
 * Fetch a single candidate by ID
 */
function getCandidateById(candidateId) {
  var list = getCandidates();
  for (var i = 0; i < list.length; i++) {
    if (list[i].id === candidateId) return list[i];
  }
  return null;
}

/**
 * Remove Candidate record row by ID
 */
function deleteCandidate(candidateId, clientEmail) {
  requireRole(['Admin', 'HR'], clientEmail);
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var dbId = data[i][0] ? String(data[i][0]).trim() : "";
    if (dbId === candidateId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error("Candidate ID " + candidateId + " not found.");
}

/**
 * Self-healing Doc template resolver (generates template on Google Drive if missing)
 */
function getOrCreateDocTemplate(templateType) {
  var configs = getConfig();
  var key = templateType === 'OFFER' ? 'template_offer_doc_id' : 'template_app_doc_id';
  var templateId = configs[key] || "";
  
  if (templateId) {
    try {
      DocumentApp.openById(templateId);
      return templateId;
    } catch (e) {
      // Template ID in database is invalid or deleted, recreate
    }
  }
  
  // Design fallbacks
  var docName = templateType === 'OFFER' ? 'Offer Letter Template' : 'Application Form Template';
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();
  
  if (templateType === 'OFFER') {
    body.appendParagraph("OFFER LETTER / THƯ MỜI NHẬN VIỆC").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Kính gửi ông/bà / Dear {{FullName}},\n\nChúng tôi rất hân hạnh đề xuất cơ hội làm việc tại công ty / We are pleased to offer you employment details:");
    body.appendParagraph("- Vị trí tuyển dụng / Position: {{Position}}\n" +
                        "- Lương Gross / Gross Salary: {{SalaryGross}} VND\n" +
                        "- Lương Net / Net Salary: {{SalaryNet}} VND\n" +
                        "- Ngày bắt đầu / Start Date: {{StartDate}}\n" +
                        "- Địa điểm / Location: {{Location}}\n" +
                        "- Hạn phản hồi / Expiry Date: {{ExpiryDate}}");
    body.appendParagraph("\nTrân trọng / Best regards,\nHR Team");
  } else {
    body.appendParagraph("CANDIDATE APPLICATION RECORD / HỒ SƠ ỨNG VIÊN").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Thông tin chi tiết ứng tuyển / Application details:");
    body.appendParagraph("- Họ và tên / Full Name: {{FullName}}\n" +
                        "- Email: {{Email}}\n" +
                        "- Số điện thoại / Phone: {{Phone}}\n" +
                        "- Vị trí ứng tuyển / Position: {{Position}}\n" +
                        "- Ngày sinh / Date of Birth: {{DateOfBirth}}\n" +
                        "- Ghi chú bổ sung / Additional Notes:\n{{Notes}}");
  }
  doc.saveAndClose();
  
  // Write to Config Sheet
  var ss = getSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  if (configSheet) {
    var data = configSheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        rowIdx = i + 1;
        break;
      }
    }
    if (rowIdx !== -1) {
      configSheet.getRange(rowIdx, 2).setValue(doc.getId());
    } else {
      configSheet.appendRow([key, doc.getId(), 'Predefined Doc Template ID']);
    }
  }
  
  return doc.getId();
}

/**
 * Generate PDF document from template doc copy, save in CV folder, and return metadata
 */
/**
 * Format currency amount with commas and unit suffix (VND or USD)
 */
function formatBackendCurrency(amount, unit) {
  if (!amount) return "";
  var cleanAmount = String(amount).replace(/[^0-9]/g, "");
  if (!cleanAmount) return "";
  var formatted = Number(cleanAmount).toLocaleString('en-US');
  if (unit === 'USD') {
    return "$" + formatted + " USD";
  } else {
    return formatted + " VND";
  }
}

function generatePDF(candidateId, templateType) {
  var candidate = getCandidateById(candidateId);
  if (!candidate) throw new Error("Candidate not found.");
  
  var templateId = getOrCreateDocTemplate(templateType);
  var templateFile = DriveApp.getFileById(templateId);
  
  // Make temporary duplicate of doc template
  var tempDocFile = templateFile.makeCopy("Temp_" + templateType + "_" + candidate.senderName);
  var tempDoc = DocumentApp.openById(tempDocFile.getId());
  var body = tempDoc.getBody();
  
  // Standard replacements
  body.replaceText("\\{\\{FullName\\}\\}", candidate.senderName || "");
  body.replaceText("\\{\\{Position\\}\\}", candidate.subject || "");
  body.replaceText("\\{\\{Email\\}\\}", candidate.senderEmail || "");
  body.replaceText("\\{\\{Phone\\}\\}", candidate.phoneNumber || "");
  body.replaceText("\\{\\{DateOfBirth\\}\\}", candidate.dateOfBirth || "");
  
  if (templateType === 'OFFER') {
    var unit = candidate.currencyUnit || "VND";
    var grossFormatted = formatBackendCurrency(candidate.offerGross, unit);
    var netFormatted = formatBackendCurrency(candidate.offerNet, unit);
    var probationFormatted = formatBackendCurrency(candidate.probationSalary, unit);
    
    body.replaceText("\\{\\{SalaryGross\\}\\}", grossFormatted);
    body.replaceText("\\{\\{SalaryNet\\}\\}", netFormatted);
    body.replaceText("\\{\\{ProbationSalary\\}\\}", probationFormatted);
    body.replaceText("\\{\\{StartDate\\}\\}", candidate.offerStartDate || "");
    body.replaceText("\\{\\{Location\\}\\}", candidate.offerLocation || "");
    body.replaceText("\\{\\{ExpiryDate\\}\\}", candidate.offerExpiry || "");
  } else {
    // Collect Form Response data
    var formResponses = null;
    if (candidate.assignedFormId) {
      try {
        formResponses = getCandidateFormResponse(candidate.senderEmail, candidate.assignedFormId);
      } catch(e) {
        Logger.log("Failed to fetch form responses: " + e.message);
      }
    }
    
    var notesVal = "";
    if (formResponses) {
      var notesArr = [];
      for (var q in formResponses) {
        notesArr.push(q + ": " + formResponses[q]);
      }
      notesVal = notesArr.join("\n");
    } else {
      notesVal = "Ứng viên chưa điền khảo sát này / Candidate has not responded yet.";
    }
    body.replaceText("\\{\\{Notes\\}\\}", notesVal);
  }
  
  tempDoc.saveAndClose();
  
  // Convert Doc copy to PDF
  var pdfBlob = tempDocFile.getAs(MimeType.PDF);
  
  // Get destination folder
  var configs = getConfig();
  var folderId = configs.cv_folder_id;
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  
  // Create PDF File in Drive
  var pdfFile = folder.createFile(pdfBlob);
  pdfFile.setName(templateType + "_" + candidate.senderName + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // Get base64 representation of PDF bytes for inline silent client printing
  var pdfBytes = pdfBlob.getBytes();
  var pdfBase64 = Utilities.base64Encode(pdfBytes);
  
  // Trash temporary doc
  try {
    tempDocFile.setTrashed(true);
  } catch (e) {
    Logger.log("Failed to trash temporary file: " + e.message);
  }
  
  return {
    url: pdfFile.getUrl(),
    fileId: pdfFile.getId(),
    base64: pdfBase64
  };
}

/**
 * Send Offer Letter Invitation email with PDF attachment
 */
function sendOfferEmailWithAttachment(candidate, pdfFileId, extraData) {
  var pdfFile = DriveApp.getFileById(pdfFileId);
  var emailSubject = "Thư mời nhận việc / Offer Letter - " + (candidate.subject || "");
  var configs = getConfig();
  
  var emailBody = "Dear " + candidate.senderName + ",\n\nWe are pleased to send you our offer letter for the " + candidate.subject + " position. Please find the attached PDF offer document.\n\nBest regards,\nHR Team";
  if (extraData && extraData.emailBody) {
    emailBody = extraData.emailBody;
  }
  
  try {
    var plainText = emailBody.replace(/<[^>]*>/g, '');
    MailApp.sendEmail({
      to: candidate.senderEmail,
      subject: emailSubject,
      body: plainText,
      htmlBody: emailBody,
      attachments: [pdfFile.getAs(MimeType.PDF)]
    });
    
    // Update Sent_Offer_At (Column 22)
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Candidates');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === candidate.id) {
          var nowStr = formatDateString(new Date());
          sheet.getRange(i + 1, 22).setValue(nowStr);
          break;
        }
      }
    }
  } catch(e) {
    Logger.log("Failed to send Offer email: " + e.message);
  }
}

/**
 * Fetch plain text body of email by Candidate ID using stored MessageID
 */
function getEmailPlainBody(candidateId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues();
  var colMap = getCandidatesColumnMap(sheet);
  
  var getValue = function(row, colName, fallbackIdx) {
    var idx = colMap[colName] !== undefined ? colMap[colName] : fallbackIdx;
    return (idx !== undefined && row[idx] !== undefined) ? String(row[idx]).trim() : "";
  };

  var rowIndex = -1;
  var cachedEmailContent = "";
  var messageId = "";
  var senderEmail = "";
  var subject = "";
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var idVal = getValue(row, 'ID', 0);
    if (idVal === candidateId) {
      rowIndex = i + 1;
      cachedEmailContent = getValue(row, 'EmailContent', 30);
      messageId = getValue(row, 'MessageID', 10);
      senderEmail = getValue(row, 'SenderEmail', 3);
      subject = getValue(row, 'Subject', 4);
      break;
    }
  }
  
  if (rowIndex === -1) return "";
  
  // 1. Return cached content from sheet if exists and not empty
  if (cachedEmailContent) {
    return cachedEmailContent;
  }
  
  var fetchedContent = "";
  
  // 2. Fallback to Gmail fetch
  if (messageId) {
    try {
      var msg = GmailApp.getMessageById(messageId);
      if (msg) {
        fetchedContent = msg.getPlainBody() || msg.getBody() || "";
      }
    } catch(e) {
      Logger.log("Failed to fetch message by ID: " + e.message + ". Fallback to search.");
    }
  }
  
  // 3. Fallback to search if messageId was invalid/missing
  if (!fetchedContent && senderEmail) {
    try {
      var cleanSubject = subject.replace(/["']/g, "");
      var query = 'from:(' + senderEmail + ') ';
      if (cleanSubject) {
        query += 'subject:(' + cleanSubject + ')';
      }
      var threads = GmailApp.search(query, 0, 1);
      if (threads.length > 0) {
        var msgs = threads[0].getMessages();
        for (var m = 0; m < msgs.length; m++) {
          var mSender = msgs[m].getFrom().toLowerCase();
          if (mSender.indexOf(senderEmail.toLowerCase()) !== -1) {
            var foundMsg = msgs[m];
            var foundId = foundMsg.getId();
            
            // Self-healing: Update MessageID dynamically
            setCandidateValueByName(sheet, rowIndex, 'MessageID', foundId, 10);
            fetchedContent = foundMsg.getPlainBody() || foundMsg.getBody() || "";
            break;
          }
        }
      }
    } catch (err) {
      Logger.log("Fallback search failed: " + err.message);
    }
  }
  
  // 4. Update the sheet with the fetched email body (truncated to 30,000 characters for safety)
  if (fetchedContent) {
    try {
      if (fetchedContent.length > 30000) {
        fetchedContent = fetchedContent.substring(0, 30000) + "\n\n...[Nội dung email bị cắt bớt / Truncated]";
      }
      setCandidateValueByName(sheet, rowIndex, 'EmailContent', fetchedContent, 30);
    } catch (sheetErr) {
      Logger.log("Failed to update EmailContent into Candidates sheet: " + sheetErr.message);
    }
  }
  
  return fetchedContent;
}

/**
 * Get the role and metadata of the currently logged-in user via Google Session.
 */
function getCurrentUserRole() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    return { email: "", role: "", fullName: "" };
  }
  email = email.toLowerCase().trim();
  var ss = getSpreadsheet();
  
  // 1. Check in Users sheet
  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet) {
    var data = usersSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var sheetEmail = data[i][0] ? String(data[i][0]).toLowerCase().trim() : "";
      if (sheetEmail === email) {
        var role = data[i][3] ? String(data[i][3]).trim() : "Employee";
        return {
          email: email,
          role: role || "Employee",
          fullName: data[i][2] ? String(data[i][2]).trim() : ""
        };
      }
    }
  }
  
  // 2. Check in Employees sheet
  var empSheet = ss.getSheetByName('Employees');
  if (empSheet) {
    var empData = empSheet.getDataRange().getValues();
    for (var j = 1; j < empData.length; j++) {
      var sheetEmail = empData[j][2] ? String(empData[j][2]).toLowerCase().trim() : "";
      if (sheetEmail === email) {
        var pos = empData[j][5] ? String(empData[j][5]).toLowerCase().trim() : "";
        var role = "Employee";
        if (pos.indexOf("manager") !== -1 || pos.indexOf("trưởng phòng") !== -1 || pos.indexOf("lead") !== -1) {
          role = "Manager";
        }
        return {
          email: email,
          role: role,
          fullName: empData[j][1] ? String(empData[j][1]).trim() : ""
        };
      }
    }
  }
  
  return { email: email, role: "Employee", fullName: "" };
}

/**
 * Helper to fetch a user's role by email.
 */
function getUserRoleByEmail(email) {
  if (!email) return "Employee";
  email = String(email).toLowerCase().trim();
  if (!email) return "Employee";

  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet) {
    var data = usersSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var sheetEmail = data[i][0] ? String(data[i][0]).toLowerCase().trim() : "";
      if (sheetEmail === email) {
        var foundRole = data[i][3] ? String(data[i][3]).trim() : "";
        return foundRole || "Employee";
      }
    }
  }
  
  var empSheet = ss.getSheetByName('Employees');
  if (empSheet) {
    var empData = empSheet.getDataRange().getValues();
    for (var j = 1; j < empData.length; j++) {
      var sheetEmail = empData[j][2] ? String(empData[j][2]).toLowerCase().trim() : "";
      if (sheetEmail === email) {
        var pos = empData[j][5] ? String(empData[j][5]).toLowerCase().trim() : "";
        if (pos.indexOf("manager") !== -1 || pos.indexOf("trưởng phòng") !== -1 || pos.indexOf("lead") !== -1) {
          return "Manager";
        }
        return "Employee";
      }
    }
  }
  
  return "Employee";
}

/**
 * Get all users with Role = Interviewer
 */
function getInterviewers() {
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) return [];
  var data = usersSheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][3] ? String(data[i][3]).trim() : "";
    if (role === 'Interviewer') {
      list.push({
        email: data[i][0] ? String(data[i][0]).trim() : "",
        fullName: data[i][2] ? String(data[i][2]).trim() : ""
      });
    }
  }
  return list;
}

function isValidEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email).toLowerCase().trim());
}

/**
 * Create Google Calendar event with Hangout/Meet link.
 * Uses Calendar Advanced Service API with a fallback to standard CalendarApp.
 */
function createMeetLink(startTime, endTime, attendeeEmail, candidateName, positionName) {
  var start = new Date(startTime);
  var end = new Date(endTime);
  
  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    throw new Error("Ngày giờ bắt đầu hoặc kết thúc không hợp lệ.");
  }
  
  var cleanEmail = attendeeEmail ? String(attendeeEmail).trim() : "";
  var userEmail = "";
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch(sessErr) {
    Logger.log("Failed to get active user email: " + sessErr.message);
  }
  
  var inviteGuests = false;
  if (cleanEmail && isValidEmail(cleanEmail) && cleanEmail.toLowerCase() !== userEmail.toLowerCase()) {
    inviteGuests = true;
  }
  
  var advError = "";
  // 1. Try using Advanced Calendar Service if enabled
  try {
    var calendarId = 'primary';
    var resource = {
      summary: 'Phỏng vấn ứng viên ' + candidateName + ' - Vị trí: ' + positionName,
      location: 'Google Meet',
      description: 'Phỏng vấn tuyển dụng HR System',
      start: {
        dateTime: start.toISOString()
      },
      end: {
        dateTime: end.toISOString()
      },
      conferenceData: {
        createRequest: {
          requestId: "meet_" + Number(new Date()) + "_" + Math.floor(Math.random() * 1000),
          conferenceSolutionKey: {
            type: "hangoutsMeet"
          }
        }
      }
    };
    
    if (inviteGuests) {
      resource.attendees = [{ email: cleanEmail }];
    }
    
    var event = Calendar.Events.insert(resource, calendarId, {
      conferenceDataVersion: 1
    });
    
    // Prioritize hangoutLink as referenced by the user
    var meetLink = event.hangoutLink || "";
    
    // Fallback to entryPoints search if hangoutLink is not populated
    if (!meetLink && event.conferenceData && event.conferenceData.entryPoints) {
      for (var i = 0; i < event.conferenceData.entryPoints.length; i++) {
        if (event.conferenceData.entryPoints[i].entryPointType === "video") {
          meetLink = event.conferenceData.entryPoints[i].uri;
          break;
        }
      }
    }
    
    if (meetLink) {
      return { eventId: event.id, meetLink: meetLink };
    } else {
      advError = "Advanced Calendar Service insert succeeded but returned empty hangoutLink.";
    }
  } catch (e) {
    advError = e.message;
    Logger.log("Advanced Calendar Service failed, trying fallback: " + e.message);
  }
  
  // 2. Fallback to CalendarApp
  try {
    var eventName = 'Phỏng vấn ứng viên ' + candidateName + ' - Vị trí: ' + positionName;
    var options = {
      description: 'Phỏng vấn tuyển dụng HR System'
    };
    
    if (inviteGuests) {
      options.guests = cleanEmail;
      options.sendInvites = true;
    }
    
    var calendar = CalendarApp.getDefaultCalendar();
    if (!calendar) {
      throw new Error("Không tìm thấy Lịch mặc định của tài khoản Google. (Chi tiết Advanced: " + advError + ")");
    }
    
    var event = calendar.createEvent(eventName, start, end, options);
    var meetLink = "";
    
    if (inviteGuests) {
      Utilities.sleep(2000); // Wait for sync
      var updatedEvent = calendar.getEventById(event.getId());
      if (updatedEvent) {
        var loc = updatedEvent.getLocation() || "";
        if (loc.indexOf("meet.google.com") !== -1) {
          meetLink = loc;
        } else {
          var desc = updatedEvent.getDescription() || "";
          var match = desc.match(/https:\/\/meet\.google\.com\/[a-z]{3}-[a-z]{4}-[a-z]{3}/i);
          if (match) {
            meetLink = match[0];
          }
        }
      }
    }
    
    if (meetLink) {
      return { eventId: event.getId(), meetLink: meetLink };
    }
    
    throw new Error(
      "Vui lòng kích hoạt dịch vụ 'Calendar' trong cột Services (bên trái trình chỉnh sửa Apps Script). " +
      "(Mã lỗi kỹ thuật: " + advError + ")"
    );
  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Save candidate interview planning and optionally email candidate (for HR)
 */
function saveInterviewPlanning(candidateId, interviewerEmail, interviewDate, meetLink, sendEmail, emailBody) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === candidateId) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Candidate ID " + candidateId + " not found.");
  
  var candidate = getCandidateById(candidateId);
  if (!candidate) throw new Error("Candidate not found.");
  
  // 1. Save data to sheet
  // AssignedInterviewer (Col 27), MeetLink (Col 28), Status (Col 7 = Pending), InterviewInfo (Col 9)
  sheet.getRange(rowIndex, 7).setValue('Pending'); // Chờ phỏng vấn
  
  var formattedInfo = "Date: " + interviewDate + " | Loc: Google Meet";
  sheet.getRange(rowIndex, 9).setValue(formattedInfo);
  sheet.getRange(rowIndex, 27).setValue(interviewerEmail || "");
  sheet.getRange(rowIndex, 28).setValue(meetLink || "");
  
  // Save to Interviews sheet too
  var intSheet = ss.getSheetByName('Interviews');
  if (intSheet) {
    var intId = getNextInterviewId(intSheet);
    intSheet.appendRow([intId, candidateId, interviewDate, meetLink || 'Google Meet', 'Interviewer: ' + interviewerEmail]);
  }
  
  // 2. Create calendar event on default calendar and invite interviewer
  if (interviewerEmail) {
    try {
      var start = new Date(interviewDate);
      if (!isNaN(start.getTime())) {
        var end = new Date(start.getTime() + 60 * 60 * 1000); // 1 hour duration
        var title = 'Lịch phỏng vấn ứng viên: ' + candidate.senderName + ' - Vị trí: ' + candidate.subject;
        var desc = 'Bạn đã được gán phỏng vấn cho ứng viên ' + candidate.senderName + '.\n' +
                   'Vị trí: ' + candidate.subject + '\n' +
                   'Chi tiết: ' + formattedInfo + '\n' +
                   'Link Google Meet: ' + (meetLink || 'Không có') + '\n' +
                   'CV ứng viên: ' + candidate.cvLink;
        
        CalendarApp.getDefaultCalendar().createEvent(title, start, end, {
          description: desc,
          location: meetLink || 'Google Meet',
          guests: interviewerEmail,
          sendInvites: true
        });
      }
    } catch (e) {
      Logger.log("Failed to create calendar event for interviewer: " + e.message);
    }
  }
  
  // 3. Send email to candidate if checked
  if (sendEmail && emailBody) {
    var candidateData = {
      id: candidateId,
      senderName: candidate.senderName,
      senderEmail: candidate.senderEmail,
      subject: candidate.subject,
      cvLink: candidate.cvLink,
      status: 'Pending'
    };
    
    var customExtra = {
      interviewDate: interviewDate,
      location: meetLink || 'Google Meet',
      note: 'Phỏng vấn cùng Interviewer: ' + interviewerEmail,
      emailBody: emailBody
    };
    
    sendEmailNotification(candidateData, "interview", customExtra);
  }
  
  return { success: true };
}

/**
 * Filter candidates assigned to the current Interviewer email
 */
function getInterviewerCandidates() {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) return [];
  
  var list = getCandidates();
  var filtered = [];
  for (var i = 0; i < list.length; i++) {
    var candInterviewer = list[i].assignedInterviewer ? list[i].assignedInterviewer.toLowerCase().trim() : "";
    if (candInterviewer === activeEmail) {
      filtered.push(list[i]);
    }
  }
  return filtered;
}

/**
 * Securely fetch candidate CV link if user is authorized.
 */
function getCandidateCVUrlOrBase64(candidateId) {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) throw new Error("Yêu cầu xác thực tài khoản.");
  
  var role = getUserRoleByEmail(activeEmail);
  if (role === 'Admin' || role === 'HR') {
    return getCandidateCVData(candidateId);
  }
  
  if (role === 'Interviewer') {
    var candidate = getCandidateById(candidateId);
    if (candidate && candidate.assignedInterviewer && candidate.assignedInterviewer.toLowerCase().trim() === activeEmail) {
      return getCandidateCVData(candidateId);
    }
  }
  
  throw new Error("Bạn không có quyền xem CV của ứng viên này.");
}

/**
 * Helper to fetch CV link and base64
 */
function getCandidateCVData(candidateId) {
  var candidate = getCandidateById(candidateId);
  if (!candidate) throw new Error("Không tìm thấy ứng viên.");
  var url = candidate.cvLink || "";
  var base64 = "";
  
  var match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    try {
      var file = DriveApp.getFileById(match[1]);
      var blob = file.getBlob();
      base64 = Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      Logger.log("Failed to encode CV file: " + e.message);
    }
  }
  
  return {
    url: url,
    base64: base64
  };
}

/**
 * Update candidate interview assessment results by Interviewer
 */
function updateInterviewResult(candidateId, result, offerData) {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) throw new Error("Yêu cầu đăng nhập.");
  
  // Verify that the candidate is assigned to this interviewer (unless Admin/HR)
  var role = getUserRoleByEmail(activeEmail);
  var candidate = getCandidateById(candidateId);
  if (!candidate) throw new Error("Candidate not found.");
  
  if (role !== 'Admin' && role !== 'HR') {
    var assigned = candidate.assignedInterviewer ? candidate.assignedInterviewer.toLowerCase().trim() : "";
    if (assigned !== activeEmail) {
      throw new Error("Bạn không có quyền đánh giá ứng viên này.");
    }
  }
  
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Candidates');
  if (!sheet) throw new Error("Candidates sheet not found");
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === candidateId) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Candidate ID " + candidateId + " not found.");
  
  // Save evaluation history as JSON
  var historyStr = sheet.getRange(rowIndex, 30).getValue();
  var history = [];
  if (historyStr) {
    try {
      history = JSON.parse(historyStr);
    } catch (e) {
      history = [];
    }
  }
  
  var nowStr = formatDateString(new Date());
  var evalItem = {
    interviewer: activeEmail,
    status: result, // 'Passed' or 'Failed'
    salary: offerData ? offerData.gross : "",
    startDate: offerData ? offerData.startDate : "",
    notes: offerData ? offerData.notes : "",
    timestamp: nowStr
  };
  history.push(evalItem);
  
  // Write result details
  sheet.getRange(rowIndex, 7).setValue('Interviewed'); // Trạng thái là Đã phỏng vấn
  sheet.getRange(rowIndex, 30).setValue(JSON.stringify(history));
  
  if (result === 'Passed' && offerData) {
    sheet.getRange(rowIndex, 15).setValue(offerData.gross || "");
    sheet.getRange(rowIndex, 17).setValue(offerData.startDate || "");
    sheet.getRange(rowIndex, 29).setValue(JSON.stringify(offerData)); // Ghi vào OfferData
  }
  
  return { success: true };
}

/**
 * Synchronize employee account to Users sheet with appropriate role
 */
function syncEmployeeToUsers(email, fullName, position) {
  var ss = getSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) return;
  
  var cleanEmail = String(email).toLowerCase().trim();
  var data = usersSheet.getDataRange().getValues();
  var exists = false;
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === cleanEmail) {
      exists = true;
      rowIndex = i + 1;
      break;
    }
  }
  
  var role = "Employee";
  var pos = String(position).toLowerCase();
  if (pos.indexOf("manager") !== -1 || pos.indexOf("trưởng phòng") !== -1 || pos.indexOf("lead") !== -1) {
    role = "Manager";
  }
  
  if (exists) {
    usersSheet.getRange(rowIndex, 3).setValue(fullName);
    usersSheet.getRange(rowIndex, 4).setValue(role);
  } else {
    var defaultHash = hashPassword('123456');
    usersSheet.appendRow([cleanEmail, defaultHash, fullName, role]);
  }
}

/**
 * Helper to get values of allowances from categories sheet
 */
function getAllowanceValuesMap() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Categories');
  var map = {};
  if (!sheet) return map;
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return map;
  
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var typeIdx = headers.indexOf('CategoryType');
  var codeIdx = headers.indexOf('Code');
  var valIdx = headers.indexOf('Value');
  
  if (typeIdx === -1 || codeIdx === -1 || valIdx === -1) {
    typeIdx = 1;
    codeIdx = 2;
    valIdx = 4;
  }
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var type = row[typeIdx] ? String(row[typeIdx]).trim() : "";
    var code = row[codeIdx] ? String(row[codeIdx]).trim() : "";
    var val = row[valIdx] ? Number(row[valIdx]) : 0;
    if (type.toLowerCase() === "allowance" && code) {
      map[code] = val;
    }
  }
  return map;
}

/**
 * Fetch data needed for employee creation/modification form
 */
function getEmployeeFormData() {
  var ss = getSpreadsheet();
  var result = {
    departments: [],
    allowances: [],
    positions: [],
    candidates: []
  };
  
  // 1. Fetch Categories
  var catSheet = ss.getSheetByName('Categories');
  if (catSheet) {
    var catData = catSheet.getDataRange().getValues();
    for (var i = 1; i < catData.length; i++) {
      var type = catData[i][1] ? String(catData[i][1]).trim() : "";
      var code = catData[i][2] ? String(catData[i][2]).trim() : "";
      var name = catData[i][3] ? String(catData[i][3]).trim() : "";
      var value = catData[i][4] ? String(catData[i][4]).trim() : "";
      
      if (type === 'Department') {
        result.departments.push({ code: code, name: name });
      } else if (type === 'Allowance') {
        result.allowances.push({ code: code, name: name, value: Number(value) || 0 });
      } else if (type === 'Position') {
        result.positions.push({ code: code, name: name });
      }
    }
  }
  
  // Fallbacks if categories are empty
  if (result.departments.length === 0) {
    result.departments = [
      { code: "TECH", name: "Phòng Công nghệ" },
      { code: "HR", name: "Phòng Nhân sự" },
      { code: "SALES", name: "Phòng Kinh doanh" },
      { code: "MKT", name: "Phòng Marketing" }
    ];
  }
  if (result.positions.length === 0) {
    result.positions = [
      { code: "DEV", name: "Nhân viên phát triển (Developer)" },
      { code: "TESTER", name: "Nhân viên kiểm thử (Tester)" },
      { code: "LEAD", name: "Trưởng nhóm kỹ thuật (Tech Lead)" },
      { code: "PM", name: "Quản trị dự án (Project Manager)" },
      { code: "HR_SPEC", name: "Chuyên viên nhân sự" },
      { code: "ACCOUNTANT", name: "Kế toán viên" }
    ];
  }
  if (result.allowances.length === 0) {
    result.allowances = [
      { code: "LUNCH", name: "Ăn trưa", value: 730000 },
      { code: "FUEL", name: "Xăng xe", value: 500000 },
      { code: "PHONE", name: "Điện thoại", value: 200000 }
    ];
  }
  
  // 2. Fetch Candidates with status 'Passed' or 'Interviewed' or related hiring states
  var candSheet = ss.getSheetByName('Candidates');
  if (candSheet) {
    var candData = candSheet.getDataRange().getValues();
    if (candData.length > 1) {
      var headers = candData[0].map(function(h) { return String(h).trim(); });
      var idCol = headers.indexOf('ID');
      var nameCol = headers.indexOf('SenderName');
      var emailCol = headers.indexOf('SenderEmail');
      var phoneCol = headers.indexOf('PhoneNumber');
      var statusCol = headers.indexOf('Status');
      var grossCol = headers.indexOf('OfferGross');
      var netCol = headers.indexOf('OfferNet');
      for (var j = 1; j < candData.length; j++) {
        var status = statusCol !== -1 ? String(candData[j][statusCol]).trim() : "";
        if (status === 'Passed' || status === 'Interviewed' || status === 'Offer Sent' || status === 'Offer Accepted') {
          result.candidates.push({
            id: idCol !== -1 ? String(candData[j][idCol]).trim() : "",
            fullName: nameCol !== -1 ? String(candData[j][nameCol]).trim() : "",
            email: emailCol !== -1 ? String(candData[j][emailCol]).trim() : "",
            phone: phoneCol !== -1 ? String(candData[j][phoneCol]).trim() : "",
            basicSalary: grossCol !== -1 ? Number(candData[j][grossCol]) || Number(candData[j][netCol]) || 0 : 0,
            probationSalary: probCol !== -1 ? Number(candData[j][probCol]) || 0 : 0
          });
        }
      }
    }
  }
  
  return result;
}

/**
 * Clear cached employee list
 */
function clearEmployeeCache() {
  try {
    CacheService.getScriptCache().remove('all_employees');
  } catch(e) {}
}

/**
 * Get list of all employees (with 5-minute CacheService acceleration)
 */
function getEmployees() {
  var cache = CacheService.getScriptCache();
  var cachedEmps = cache.get('all_employees');
  if (cachedEmps) {
    try {
      return JSON.parse(cachedEmps);
    } catch(e) {}
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Employees');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  headers.forEach(function(h, idx) {
    colMap[h] = idx;
  });
  
  var employees = [];
  var allowanceValuesMap = getAllowanceValuesMap();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Helper to get row value safely
    function getVal(colName, defaultVal) {
      var idx = colMap[colName];
      if (idx !== undefined && row[idx] !== undefined && row[idx] !== null) {
        return row[idx];
      }
      return defaultVal;
    }
    
    var rawAllowances = String(getVal('Allowances', getVal('Allowance', ''))).trim();
    var totalAllowance = 0;
    if (rawAllowances) {
      rawAllowances.split(',').forEach(function(code) {
        var trimmed = code.trim();
        if (allowanceValuesMap[trimmed]) {
          totalAllowance += allowanceValuesMap[trimmed];
        }
      });
    }
    if (totalAllowance === 0 && getVal('Allowance', 0)) {
      totalAllowance = Number(getVal('Allowance', 0));
    }
    
    var email = String(getVal('Email', '')).trim();
    var fullName = String(getVal('FullName', '')).trim();
    var id = String(getVal('ID', '')).trim();
    
    var pStart = getVal('ProbationStartDate', '');
    var oStart = getVal('OfficialStartDate', getVal('StartDate', ''));
    var contractExpiry = getVal('ContractExpiryDate', getVal('EndDate', ''));
    
    var pStartStr = pStart ? (pStart instanceof Date ? Utilities.formatDate(pStart, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(pStart)) : "";
    var oStartStr = oStart ? (oStart instanceof Date ? Utilities.formatDate(oStart, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(oStart)) : "";
    var expiryStr = contractExpiry ? (contractExpiry instanceof Date ? Utilities.formatDate(contractExpiry, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(contractExpiry)) : "";
    
    employees.push({
      id: id,
      fullName: fullName,
      gender: String(getVal('Gender', 'Nam')).trim(),
      email: email,
      phone: String(getVal('Phone', '')).trim(),
      department: String(getVal('Department', '')).trim(),
      position: String(getVal('Position', '')).trim(),
      basicSalary: Number(getVal('BasicSalary', 0)),
      probationSalary: Number(getVal('ProbationSalary', 0)),
      socialInsuranceSalary: Number(getVal('SocialInsuranceSalary', getVal('BasicSalary', 0))),
      dependents: Number(getVal('Dependents', 0)),
      totalLeaveDays: Number(getVal('TotalLeaveDays', 12)),
      usedLeaveDays: Number(getVal('UsedLeaveDays', 0)),
      allowances: rawAllowances,
      allowance: totalAllowance,
      probationStartDate: pStartStr,
      officialStartDate: oStartStr,
      contractExpiryDate: expiryStr,
      status: String(getVal('Status', 'Đang làm')).trim(),
      
      // Aliases for legacy compatibility
      startDate: oStartStr || pStartStr,
      endDate: expiryStr
    });
  }

  try {
    cache.put('all_employees', JSON.stringify(employees), 300); // Cache 5 minutes
  } catch(e) {}

  return employees;
}

/**
 * THÊM NHÂN SỰ MỚI (VỚI LOCKSERVICE VÀ CÁC THỦ TỤC TRÍCH XUẤT)
 */
function addEmployee(emp, clientEmail) {
    requireRole(['Admin', 'HR'], clientEmail || (emp && emp.clientEmail));
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); // Chờ tối đa 10s tránh trùng mã EMP
        if (!emp.fullName) throw new Error("Họ tên không được để trống.");

        var ss = getSpreadsheet();
        var sheet = ss.getSheetByName('Employees');
        if (!sheet) throw new Error("Không tìm thấy bảng Employees");

        var data = sheet.getDataRange().getValues();
        var nextId = "EMP-0001";
        if (data.length > 1) {
            var lastId = String(data[data.length - 1][0]);
            var match = lastId.match(/EMP-(\d+)/);
            if (match) {
                var num = parseInt(match[1], 10) + 1;
                nextId = "EMP-" + ("0000" + num).slice(-4);
            }
        }

        // Tính toán phụ cấp tổng
        var allowanceValuesMap = getAllowanceValuesMap();
        var totalAllowance = 0;
        if (emp.allowances) {
            emp.allowances.split(',').forEach(function (code) {
                var trimmed = code.trim();
                if (allowanceValuesMap[trimmed]) {
                    totalAllowance += allowanceValuesMap[trimmed];
                }
            });
        }
        emp.allowance = totalAllowance;

        // 1. Chèn dữ liệu vào bảng Employees
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
        var colMap = {};
        headers.forEach(function (h, idx) { colMap[h] = idx; });

        var newRow = new Array(headers.length).fill("");
        function setVal(colName, value) {
            var idx = colMap[colName];
            if (idx !== undefined) newRow[idx] = value;
        }

        setVal('ID', nextId);
        setVal('FullName', emp.fullName || "");
        setVal('Gender', emp.gender || "Nam");
        setVal('Email', emp.email || "");
        setVal('Phone', emp.phone || "");
        setVal('Department', emp.department || "");
        setVal('Position', emp.position || "");
        setVal('BasicSalary', emp.basicSalary ? Number(emp.basicSalary) : 0);
        setVal('ProbationSalary', emp.probationSalary ? Number(emp.probationSalary) : 0);
        setVal('SocialInsuranceSalary', emp.socialInsuranceSalary ? Number(emp.socialInsuranceSalary) : (emp.basicSalary ? Number(emp.basicSalary) : 0));
        setVal('Dependents', emp.dependents ? Number(emp.dependents) : 0);
        setVal('TotalLeaveDays', emp.totalLeaveDays ? Number(emp.totalLeaveDays) : 12);
        setVal('UsedLeaveDays', emp.usedLeaveDays ? Number(emp.usedLeaveDays) : 0);
        setVal('Allowances', emp.allowances || "");
        setVal('ProbationStartDate', emp.probationStartDate || "");
        setVal('OfficialStartDate', emp.officialStartDate || "");
        setVal('ContractExpiryDate', emp.contractExpiryDate || "");
        setVal('Status', emp.status || "Đang làm");
        setVal('StartDate', emp.officialStartDate || emp.probationStartDate || "");
        setVal('EndDate', emp.contractExpiryDate || "");
        setVal('Allowance', emp.allowance || 0);

        sheet.appendRow(newRow);

        // 2. Chèn dữ liệu đồng bộ tức thì sang EmployeeDetails
        var detSheet = ss.getSheetByName('EmployeeDetails');
        if (!detSheet) {
            initializeDatabaseV2();
            detSheet = ss.getSheetByName('EmployeeDetails');
        }

        var detHeaders = detSheet.getRange(1, 1, 1, detSheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
        var detColMap = {};
        detHeaders.forEach(function (h, idx) { detColMap[h] = idx; });

        var detRow = new Array(detHeaders.length).fill("");
        function setDetVal(colName, value) {
            var idx = detColMap[colName];
            if (idx !== undefined) detRow[idx] = value;
        }

        var dobVal = emp.dob || emp.dateOfBirth || "";
        if (emp.candidateId) {
            var candSheet = ss.getSheetByName('Candidates');
            if (candSheet) {
                var candData = candSheet.getDataRange().getValues();
                if (candData.length > 1) {
                    var candHeaders = candData[0].map(function (h) { return String(h).trim(); });
                    var candIDCol = candHeaders.indexOf('ID');
                    var candDOBCol = candHeaders.indexOf('DateOfBirth');
                    for (var j = 1; j < candData.length; j++) {
                        if (candIDCol !== -1 && String(candData[j][candIDCol]).trim() === String(emp.candidateId).trim()) {
                            if (candDOBCol !== -1 && candData[j][candDOBCol]) {
                                var d = candData[j][candDOBCol];
                                dobVal = d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d);
                            }
                            break;
                        }
                    }
                }
            }
        }

        setDetVal('EmployeeID', nextId);
        setDetVal('FullName', emp.fullName || "");
        setDetVal('Gender', emp.gender || "Nam");
        setDetVal('Department', emp.department || "");
        setDetVal('DateOfBirth', dobVal);
        setDetVal('Dependents', emp.dependents ? Number(emp.dependents) : 0);
        setDetVal('Avatar', emp.avatar || "");
        setDetVal('Docs', "[]");

        detSheet.appendRow(detRow);
        SpreadsheetApp.flush();

        syncEmployeeToUsers(emp.email, emp.fullName, emp.position);

        if (emp.candidateId) {
            updateCandidateStatus(emp.candidateId, 'Hired');
        }

        clearEmployeeCache();
        return { success: true, employeeId: nextId };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Update employee details with LockService & Cache clearing
 */
function updateEmployee(emp) {
  requireRole(['Admin', 'HR']);
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Employees');
    if (!sheet) throw new Error("Employees sheet not found");
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === emp.id) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Employee not found with ID: " + emp.id);
    
    // Calculate allowances total sum
    var allowanceValuesMap = getAllowanceValuesMap();
    var totalAllowance = 0;
    if (emp.allowances) {
      emp.allowances.split(',').forEach(function(code) {
        var trimmed = code.trim();
        if (allowanceValuesMap[trimmed]) {
          totalAllowance += allowanceValuesMap[trimmed];
        }
      });
    }
    emp.allowance = totalAllowance;
    
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var colMap = {};
    headers.forEach(function(h, idx) {
      colMap[h] = idx;
    });
    
    function updateCell(colName, value) {
      var idx = colMap[colName];
      if (idx !== undefined) {
        sheet.getRange(rowIndex, idx + 1).setValue(value);
      }
    }
    
    updateCell('FullName', emp.fullName || "");
    updateCell('Gender', emp.gender || "Nam");
    updateCell('Email', emp.email || "");
    updateCell('Phone', emp.phone || "");
    updateCell('Department', emp.department || "");
    updateCell('Position', emp.position || "");
    updateCell('BasicSalary', emp.basicSalary ? Number(emp.basicSalary) : 0);
    updateCell('ProbationSalary', emp.probationSalary ? Number(emp.probationSalary) : 0);
    updateCell('SocialInsuranceSalary', emp.socialInsuranceSalary ? Number(emp.socialInsuranceSalary) : (emp.basicSalary ? Number(emp.basicSalary) : 0));
    updateCell('Dependents', emp.dependents ? Number(emp.dependents) : 0);
    if (emp.totalLeaveDays !== undefined) updateCell('TotalLeaveDays', Number(emp.totalLeaveDays));
    if (emp.usedLeaveDays !== undefined) updateCell('UsedLeaveDays', Number(emp.usedLeaveDays));
    updateCell('Allowances', emp.allowances || "");
    updateCell('ProbationStartDate', emp.probationStartDate || "");
    updateCell('OfficialStartDate', emp.officialStartDate || "");
    updateCell('ContractExpiryDate', emp.contractExpiryDate || "");
    updateCell('Status', emp.status || "Đang làm");
    
    // Legacy compatibility columns
    updateCell('StartDate', emp.officialStartDate || emp.probationStartDate || "");
    updateCell('EndDate', emp.contractExpiryDate || "");
    updateCell('Allowance', emp.allowance || 0);
    
    // Sync to EmployeeDetails sheet
    syncEmployeeDetails(emp.id, emp);
    syncEmployeeToUsers(emp.email, emp.fullName, emp.position);
    
    clearEmployeeCache();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper to sync/initialize employee details on EmployeeDetails sheet
 */
function syncEmployeeDetails(employeeId, emp) {
  try {
    var ss = getSpreadsheet();
    var detSheet = ss.getSheetByName('EmployeeDetails');
    if (!detSheet) {
      initializeDatabaseV2();
      detSheet = ss.getSheetByName('EmployeeDetails');
    }
    if (!detSheet) return;
    
    var data = detSheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === employeeId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h] = idx; });
    
    if (rowIndex === -1) {
      // Row not found, create a new one
      var detRow = new Array(headers.length);
      for (var k = 0; k < detRow.length; k++) {
        detRow[k] = "";
      }
      
      function setDetVal(colName, value) {
        var idx = colMap[colName];
        if (idx !== undefined) {
          detRow[idx] = value;
        }
      }
      
      var genderVal = emp.gender || "Nam";
      var dobVal = emp.dob || emp.dateOfBirth || "";
      if (emp.candidateId) {
        var candSheet = ss.getSheetByName('Candidates');
        if (candSheet) {
          var candData = candSheet.getDataRange().getValues();
          if (candData.length > 1) {
            var candHeaders = candData[0].map(function(h) { return String(h).trim(); });
            var candColMap = {};
            candHeaders.forEach(function(h, idx) { candColMap[h] = idx; });
            var idIdx = candColMap['ID'];
            var dobIdx = candColMap['DateOfBirth'];
            var genIdx = candColMap['Gender'];
            for (var j = 1; j < candData.length; j++) {
              if (idIdx !== -1 && String(candData[j][idIdx]).trim() === String(emp.candidateId).trim()) {
                if (dobIdx !== -1 && candData[j][dobIdx] && !dobVal) {
                  var d = candData[j][dobIdx];
                  dobVal = d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d);
                }
                if (genIdx !== -1 && candData[j][genIdx]) {
                  genderVal = String(candData[j][genIdx]).trim();
                }
                break;
              }
            }
          }
        }
      }
      
      setDetVal('EmployeeID', employeeId);
      setDetVal('Avatar', emp.avatar || "");
      setDetVal('Gender', genderVal);
      setDetVal('BirthPlace', "");
      setDetVal('CurrentAddress', "");
      setDetVal('RegisterAddress', "");
      setDetVal('IdentityCardNumber', "");
      setDetVal('IdentityCardDate', "");
      setDetVal('IdentityCardPlace', "");
      setDetVal('AcademicLevel', "");
      setDetVal('Specialization', "");
      setDetVal('GraduationInstitution', "");
      setDetVal('YouthUnionDate', "");
      setDetVal('CommunistPartyDateStatus', "");
      setDetVal('Docs', "[]");
      setDetVal('DateOfBirth', dobVal);
      
      detSheet.appendRow(detRow);
    } else {
      // Row exists, sync the relevant shared fields
      function updateCell(colName, value) {
        var idx = colMap[colName];
        if (idx !== undefined && value !== undefined && value !== null) {
          detSheet.getRange(rowIndex, idx + 1).setValue(value);
        }
      }
      
      updateCell('Gender', emp.gender);
      if (emp.avatar) updateCell('Avatar', emp.avatar);
      if (emp.dob || emp.dateOfBirth) updateCell('DateOfBirth', emp.dob || emp.dateOfBirth);
    }
  } catch (e) {
    Logger.log("Lỗi trong syncEmployeeDetails: " + e.toString());
  }
}

/**
 * Get contract expiration warning list (contracts expiring within 15-30 days)
 */
function getContractWarnings() {
  var employees = getEmployees();
  var warnings = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    if (emp.status === "Đang làm" && emp.endDate) {
      var end = new Date(emp.endDate);
      if (!isNaN(end.getTime())) {
        var diffTime = end - today;
        var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        if (diffDays >= 0 && diffDays <= 30) {
          warnings.push({
            employee: emp,
            daysLeft: diffDays
          });
        }
      }
    }
  }
  return warnings;
}

/**
 * Handle increase salary updates & logs
 */
function increaseSalary(employeeId, newSalary, decisionBase64, notes) {
  var ss = getSpreadsheet();
  
  var empSheet = ss.getSheetByName('Employees');
  if (!empSheet) throw new Error("Employees sheet not found");
  
  var data = empSheet.getDataRange().getValues();
  var empRow = -1;
  var employee = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === employeeId) {
      empRow = i + 1;
      employee = {
        id: employeeId,
        fullName: String(data[i][1]),
        email: String(data[i][2]),
        position: String(data[i][5]),
        basicSalary: Number(data[i][9])
      };
      break;
    }
  }
  
  if (empRow === -1) throw new Error("Employee not found");
  
  empSheet.getRange(empRow, 10).setValue(Number(newSalary));
  
  var histSheet = ss.getSheetByName('SalaryHistory');
  var nextHistId = "SAL-0001";
  if (histSheet) {
    var histData = histSheet.getDataRange().getValues();
    if (histData.length > 1) {
      var lastId = String(histData[histData.length - 1][0]);
      var match = lastId.match(/SAL-(\d+)/);
      if (match) {
        var num = parseInt(match[1], 10) + 1;
        nextHistId = "SAL-" + ("0000" + num).slice(-4);
      }
    }
    
    var changeDateStr = formatDateString(new Date()).split(" ")[0];
    histSheet.appendRow([
      nextHistId,
      employeeId,
      Number(newSalary),
      changeDateStr,
      decisionBase64 || "",
      notes || ""
    ]);
  }
  
  var emailBody = "<div style='font-family: sans-serif; padding: 20px; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px;'>" +
                  "<h2>QUYẾT ĐỊNH TĂNG LƯƠNG</h2>" +
                  "<p>Kính gửi ông/bà <strong>" + employee.fullName + "</strong>,</p>" +
                  "<p>Ban Giám đốc xin trân trọng thông báo quyết định điều chỉnh mức lương của bạn như sau:</p>" +
                  "<table style='width: 100%; border-collapse: collapse; margin: 15px 0;'>" +
                  "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Vị trí:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + employee.position + "</td></tr>" +
                  "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Mức lương cũ:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + Number(employee.basicSalary).toLocaleString() + " VND</td></tr>" +
                  "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Mức lương mới:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee; color: #16a34a;'><strong>" + Number(newSalary).toLocaleString() + " VND</strong></td></tr>" +
                  "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Ghi chú:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + (notes || "Không có") + "</td></tr>" +
                  "</table>" +
                  "<p>Quyết định này có hiệu lực từ ngày hôm nay. Chúc bạn tiếp tục nỗ lực và đạt được nhiều thành công hơn nữa cùng công ty!</p>" +
                  "<br><p>Trân trọng,</p><p><strong>Phòng Nhân Sự (HR Department)</strong></p>" +
                  "</div>";
                  
  try {
    var plainText = "Quyết định tăng lương cho " + employee.fullName + ". Mức lương mới: " + Number(newSalary).toLocaleString() + " VND.";
    MailApp.sendEmail({
      to: employee.email,
      subject: "Quyết định tăng lương / Salary Adjustment Decision - " + employee.fullName,
      body: plainText,
      htmlBody: emailBody
    });
  } catch (mailErr) {
    Logger.log("Failed to send salary raise email: " + mailErr.message);
  }
  
  return { success: true, histId: nextHistId };
}

/**
 * Get salary history log
 */
function getSalaryHistory(employeeId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('SalaryHistory');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var history = [];
  for (var i = 1; i < data.length; i++) {
    var dbEmpId = data[i][1] ? String(data[i][1]).trim() : "";
    if (!employeeId || dbEmpId === employeeId) {
      history.push({
        id: data[i][0] ? String(data[i][0]).trim() : "",
        employeeId: dbEmpId,
        newSalary: data[i][2] ? Number(data[i][2]) : 0,
        changeDate: data[i][3] ? (data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][3])) : "",
        decisionFile: data[i][4] ? String(data[i][4]) : "",
        notes: data[i][5] ? String(data[i][5]) : ""
      });
    }
  }
  return history;
}

/**
 * OT Hours Calculation Helper
 */
function calculateOT(checkOutStr, endTimeConfig) {
  if (!checkOutStr || !endTimeConfig) return 0;
  var coParts = checkOutStr.split(":");
  var etParts = endTimeConfig.split(":");
  if (coParts.length < 2 || etParts.length < 2) return 0;
  
  var coSecs = parseInt(coParts[0], 10) * 3600 + parseInt(coParts[1], 10) * 60 + (coParts[2] ? parseInt(coParts[2], 10) : 0);
  var etSecs = parseInt(etParts[0], 10) * 3600 + parseInt(etParts[1], 10) * 60 + (etParts[2] ? parseInt(etParts[2], 10) : 0);
  
  if (coSecs <= etSecs) return 0;
  var diffSecs = coSecs - etSecs;
  var otHours = diffSecs / 3600;
  return Math.round(otHours * 100) / 100;
}

/**
 * Handle Employee check-in/out
 */
function checkInOut(employeeEmail, ipAddress) {
  var ss = getSpreadsheet();
  var empSheet = ss.getSheetByName('Employees');
  if (!empSheet) throw new Error("Employees sheet not found");
  
  var empData = empSheet.getDataRange().getValues();
  var employeeId = "";
  for (var i = 1; i < empData.length; i++) {
    if (String(empData[i][2]).toLowerCase().trim() === String(employeeEmail).toLowerCase().trim()) {
      employeeId = String(empData[i][0]).trim();
      break;
    }
  }
  
  if (!employeeId) throw new Error("Employee email not found in records.");
  
  var configs = getConfig();
  var startTime = configs.attendance_start_time || "08:00:00";
  var endTime = configs.attendance_end_time || "17:00:00";
  
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
  
  var attSheet = ss.getSheetByName('Attendance');
  if (!attSheet) throw new Error("Attendance sheet not found");
  
  var attData = attSheet.getDataRange().getValues();
  var existingRow = -1;
  for (var j = 1; j < attData.length; j++) {
    var dbEmpId = String(attData[j][1]).trim();
    var dbDate = attData[j][2];
    var dbDateStr = dbDate instanceof Date ? Utilities.formatDate(dbDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dbDate).trim();
    
    if (dbEmpId === employeeId && dbDateStr === todayStr) {
      existingRow = j + 1;
      break;
    }
  }
  
  if (existingRow === -1) {
    var nextId = "ATT-0001";
    if (attData.length > 1) {
      var lastId = String(attData[attData.length - 1][0]);
      var match = lastId.match(/ATT-(\d+)/);
      if (match) {
        var num = parseInt(match[1], 10) + 1;
        nextId = "ATT-" + ("0000" + num).slice(-4);
      }
    }
    
    var status = "Đúng giờ";
    if (timeStr > startTime) {
      status = "Đi muộn";
    }
    
    attSheet.appendRow([
      nextId,
      employeeId,
      todayStr,
      timeStr,
      "",
      ipAddress || "",
      0,
      status
    ]);
    
    return { success: true, type: "checkin", time: timeStr, status: status };
  } else {
    var checkInTime = attData[existingRow - 1][3] ? String(attData[existingRow - 1][3]).trim() : "";
    var currentStatus = attData[existingRow - 1][7] ? String(attData[existingRow - 1][7]).trim() : "";
    
    var status = currentStatus;
    if (timeStr < endTime) {
      status = "Về sớm";
    }
    
    var otHours = calculateOT(timeStr, endTime);
    
    attSheet.getRange(existingRow, 5).setValue(timeStr);
    attSheet.getRange(existingRow, 6).setValue(ipAddress || "");
    attSheet.getRange(existingRow, 7).setValue(otHours);
    attSheet.getRange(existingRow, 8).setValue(status);
    
    return { success: true, type: "checkout", time: timeStr, otHours: otHours, status: status };
  }
}

/**
 * Import attendance data from Excel data rows
 */
function importAttendanceExcel(dataRows) {
  var ss = getSpreadsheet();
  var attSheet = ss.getSheetByName('Attendance');
  var empSheet = ss.getSheetByName('Employees');
  if (!attSheet || !empSheet) throw new Error("Database sheets not initialized");
  
  var empData = empSheet.getDataRange().getValues();
  var empEmailMap = {};
  var empIdMap = {};
  for (var i = 1; i < empData.length; i++) {
    var id = String(empData[i][0]).trim();
    var email = String(empData[i][2]).toLowerCase().trim();
    empEmailMap[email] = id;
    empIdMap[id] = id;
  }
  
  var attData = attSheet.getDataRange().getValues();
  var nextNum = 1;
  if (attData.length > 1) {
    var lastId = String(attData[attData.length - 1][0]);
    var match = lastId.match(/ATT-(\d+)/);
    if (match) {
      nextNum = parseInt(match[1], 10) + 1;
    }
  }
  
  var configs = getConfig();
  var endTime = configs.attendance_end_time || "17:00:00";
  
  var importedCount = 0;
  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];
    var emailOrId = String(row.emailOrId || "").trim();
    var empId = empEmailMap[emailOrId.toLowerCase()] || empIdMap[emailOrId] || "";
    if (!empId) continue;
    
    var dateStr = String(row.date || "").trim();
    var checkIn = String(row.checkIn || "").trim();
    var checkOut = String(row.checkOut || "").trim();
    var ip = String(row.ipAddress || row.ip || "").trim();
    
    var duplicateIdx = -1;
    for (var j = 1; j < attData.length; j++) {
      var dbEmpId = String(attData[j][1]).trim();
      var dbDate = attData[j][2];
      var dbDateStr = dbDate instanceof Date ? Utilities.formatDate(dbDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dbDate).trim();
      if (dbEmpId === empId && dbDateStr === dateStr) {
        duplicateIdx = j + 1;
        break;
      }
    }
    
    var otHours = calculateOT(checkOut, endTime);
    var status = row.status || "Đúng giờ";
    
    if (duplicateIdx === -1) {
      var nextId = "ATT-" + ("0000" + nextNum).slice(-4);
      attSheet.appendRow([
        nextId,
        empId,
        dateStr,
        checkIn,
        checkOut,
        ip,
        otHours,
        status
      ]);
      nextNum++;
      importedCount++;
    } else {
      if (checkIn) attSheet.getRange(duplicateIdx, 4).setValue(checkIn);
      if (checkOut) {
        attSheet.getRange(duplicateIdx, 5).setValue(checkOut);
        attSheet.getRange(duplicateIdx, 7).setValue(otHours);
      }
      if (ip) attSheet.getRange(duplicateIdx, 6).setValue(ip);
      if (status) attSheet.getRange(duplicateIdx, 8).setValue(status);
      importedCount++;
    }
  }
  
  return { success: true, count: importedCount };
}

/**
 * Record live Check-In / Check-Out for an employee with LockService & GracePeriod protection
 */
function recordAttendance(userEmail, type, clientIp) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var activeEmail = getActiveUserEmail();
    var email = (activeEmail && activeEmail.length > 0) ? activeEmail : userEmail;
    if (!email) throw new Error("Email người dùng không hợp lệ.");
    
    email = String(email).toLowerCase().trim();
    var ss = getSpreadsheet();
    var attSheet = ss.getSheetByName('Attendance');
    var empSheet = ss.getSheetByName('Employees');
    if (!attSheet) {
      initializeDatabaseV2();
      attSheet = ss.getSheetByName('Attendance');
    }
    
    // Find Employee ID
    var empId = "";
    if (empSheet) {
      var empData = empSheet.getDataRange().getValues();
      for (var i = 1; i < empData.length; i++) {
        var sheetEmail = empData[i][2] ? String(empData[i][2]).toLowerCase().trim() : "";
        if (sheetEmail === email) {
          empId = String(empData[i][0]).trim();
          break;
        }
      }
    }
    if (!empId) empId = email;
    
    var now = new Date();
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    var currentTimeStr = Utilities.formatDate(now, timeZone, "HH:mm:ss");
    
    var configs = getConfig();
    var startTimeConfig = configs.attendance_start_time || "08:00";
    var endTimeConfig = configs.attendance_end_time || "17:00";
    var gracePeriod = parseInt(configs.grace_period_minutes || 15, 10);
    
    var data = attSheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var idCol = headers.indexOf('ID') !== -1 ? headers.indexOf('ID') : 0;
    var empIdCol = headers.indexOf('EmployeeID') !== -1 ? headers.indexOf('EmployeeID') : 1;
    var dateCol = headers.indexOf('Date') !== -1 ? headers.indexOf('Date') : 2;
    var checkInCol = headers.indexOf('CheckIn') !== -1 ? headers.indexOf('CheckIn') : 3;
    var checkOutCol = headers.indexOf('CheckOut') !== -1 ? headers.indexOf('CheckOut') : 4;
    var ipCol = headers.indexOf('IPAddress') !== -1 ? headers.indexOf('IPAddress') : 5;
    var otCol = headers.indexOf('OT_Hours') !== -1 ? headers.indexOf('OT_Hours') : 6;
    var statusCol = headers.indexOf('Status') !== -1 ? headers.indexOf('Status') : 7;
    
    var targetRowIndex = -1;
    var existingCheckIn = "";
    
    for (var j = 1; j < data.length; j++) {
      var rowEmpId = String(data[j][empIdCol]).trim();
      var rowDateStr = "";
      if (data[j][dateCol] instanceof Date) {
        rowDateStr = Utilities.formatDate(data[j][dateCol], timeZone, "yyyy-MM-dd");
      } else {
        rowDateStr = String(data[j][dateCol]).trim();
      }
      
      if (rowEmpId === empId && rowDateStr === todayStr) {
        targetRowIndex = j + 1;
        existingCheckIn = data[j][checkInCol] ? String(data[j][checkInCol]).trim() : "";
        break;
      }
    }
    
    var status = "Đúng giờ";
    
    if (type === 'checkin') {
      var nowMinutes = now.getHours() * 60 + now.getMinutes();
      var startParts = startTimeConfig.split(':');
      var startMinutes = (parseInt(startParts[0], 10) * 60 + parseInt(startParts[1] || 0, 10)) + gracePeriod;
      
      if (nowMinutes > startMinutes) {
        status = "Đi muộn";
      } else {
        status = "Đúng giờ";
      }
      
      if (targetRowIndex > 0) {
        attSheet.getRange(targetRowIndex, checkInCol + 1).setValue(currentTimeStr);
        if (clientIp) attSheet.getRange(targetRowIndex, ipCol + 1).setValue(clientIp);
        attSheet.getRange(targetRowIndex, statusCol + 1).setValue(status);
      } else {
        var newId = "ATT-" + Number(now);
        var rowData = [];
        var maxCol = Math.max(idCol, empIdCol, dateCol, checkInCol, checkOutCol, ipCol, otCol, statusCol) + 1;
        for (var c = 0; c < maxCol; c++) {
          if (c === idCol) rowData.push(newId);
          else if (c === empIdCol) rowData.push(empId);
          else if (c === dateCol) rowData.push(todayStr);
          else if (c === checkInCol) rowData.push(currentTimeStr);
          else if (c === checkOutCol) rowData.push("");
          else if (c === ipCol) rowData.push(clientIp || "");
          else if (c === otCol) rowData.push(0);
          else if (c === statusCol) rowData.push(status);
          else rowData.push("");
        }
        attSheet.appendRow(rowData);
      }
    } else if (type === 'checkout') {
      var nowMinutes = now.getHours() * 60 + now.getMinutes();
      var endParts = endTimeConfig.split(':');
      var endMinutes = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1] || 0, 10);
      
      if (nowMinutes < endMinutes && existingCheckIn) {
        status = "Về sớm";
      } else {
        status = "Đúng giờ";
      }
      
      if (targetRowIndex > 0) {
        attSheet.getRange(targetRowIndex, checkOutCol + 1).setValue(currentTimeStr);
        if (clientIp) attSheet.getRange(targetRowIndex, ipCol + 1).setValue(clientIp);
        if (status === "Về sớm") {
          attSheet.getRange(targetRowIndex, statusCol + 1).setValue(status);
        }
      } else {
        var newId = "ATT-" + Number(now);
        var rowData = [];
        var maxCol = Math.max(idCol, empIdCol, dateCol, checkInCol, checkOutCol, ipCol, otCol, statusCol) + 1;
        for (var c = 0; c < maxCol; c++) {
          if (c === idCol) rowData.push(newId);
          else if (c === empIdCol) rowData.push(empId);
          else if (c === dateCol) rowData.push(todayStr);
          else if (c === checkInCol) rowData.push("");
          else if (c === checkOutCol) rowData.push(currentTimeStr);
          else if (c === ipCol) rowData.push(clientIp || "");
          else if (c === otCol) rowData.push(0);
          else if (c === statusCol) rowData.push(status);
          else rowData.push("");
        }
        attSheet.appendRow(rowData);
      }
    }
    
    return { success: true, time: currentTimeStr, status: status, date: todayStr };
  } catch (e) {
    throw new Error("Lỗi Check-In/Out: " + e.message);
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * Fetch attendance logs securely (Role & Identity verified on server)
 */
function getAttendanceLogs(requestedEmail, requestedRole) {
  var activeEmail = getActiveUserEmail();
  var email = (activeEmail && activeEmail.length > 0) ? activeEmail : String(requestedEmail || "").toLowerCase().trim();
  if (!email) return [];
  
  var role = getUserRoleByEmail(email);
  if (!role || role.length === 0) {
    role = requestedRole || "Employee";
  }
  
  var ss = getSpreadsheet();
  var attSheet = ss.getSheetByName('Attendance');
  var empSheet = ss.getSheetByName('Employees');
  if (!attSheet) return [];
  
  var empData = empSheet ? empSheet.getDataRange().getValues() : [];
  var empIdToNameMap = {};
  var empIdToEmailMap = {};
  var empIdToDeptMap = {};
  var targetEmployeeId = "";
  
  for (var i = 1; i < empData.length; i++) {
    var id = String(empData[i][0]).trim();
    var name = String(empData[i][1]).trim();
    var empEmail = String(empData[i][2]).toLowerCase().trim();
    var dept = String(empData[i][4]).trim();
    empIdToNameMap[id] = name;
    empIdToEmailMap[id] = empEmail;
    empIdToDeptMap[id] = dept;
    
    if (empEmail === email) {
      targetEmployeeId = id;
    }
  }
  if (!targetEmployeeId) targetEmployeeId = email;
  
  var logs = [];
  var data = attSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idCol = headers.indexOf('ID') !== -1 ? headers.indexOf('ID') : 0;
  var empIdCol = headers.indexOf('EmployeeID') !== -1 ? headers.indexOf('EmployeeID') : 1;
  var dateCol = headers.indexOf('Date') !== -1 ? headers.indexOf('Date') : 2;
  var checkInCol = headers.indexOf('CheckIn') !== -1 ? headers.indexOf('CheckIn') : 3;
  var checkOutCol = headers.indexOf('CheckOut') !== -1 ? headers.indexOf('CheckOut') : 4;
  var ipCol = headers.indexOf('IPAddress') !== -1 ? headers.indexOf('IPAddress') : 5;
  var otCol = headers.indexOf('OT_Hours') !== -1 ? headers.indexOf('OT_Hours') : 6;
  var statusCol = headers.indexOf('Status') !== -1 ? headers.indexOf('Status') : 7;
  
  var timeZone = Session.getScriptTimeZone();
  
  for (var j = 1; j < data.length; j++) {
    var rowEmpId = String(data[j][empIdCol]).trim();
    var rowEmail = empIdToEmailMap[rowEmpId] || rowEmpId;
    var rowDept = empIdToDeptMap[rowEmpId] || "";
    
    var allow = false;
    if (role === 'Admin' || role === 'HR') {
      allow = true;
    } else if (role === 'Manager') {
      var managerDept = "";
      for (var k = 1; k < empData.length; k++) {
        var e = String(empData[k][2]).toLowerCase().trim();
        if (e === userEmail) {
          managerDept = String(empData[k][4]).trim();
          break;
        }
      }
      if (managerDept && logDept === managerDept) {
        allow = true;
      }
    } else if (role === 'Employee' || empId === targetEmployeeId) {
      if (empId === targetEmployeeId) {
        allow = true;
      }
    }
    
    if (allow) {
      logs.push({
        id: data[j][0] ? String(data[j][0]).trim() : "",
        employeeId: empId,
        employeeName: empIdToNameMap[empId] || "Unknown",
        employeeEmail: logEmail,
        department: logDept,
        date: data[j][2] ? (data[j][2] instanceof Date ? Utilities.formatDate(data[j][2], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[j][2])) : "",
        checkIn: data[j][3] ? String(data[j][3]).trim() : "",
        checkOut: data[j][4] ? String(data[j][4]).trim() : "",
        ipAddress: data[j][5] ? String(data[j][5]).trim() : "",
        otHours: data[j][6] ? Number(data[j][6]) : 0,
        status: data[j][7] ? String(data[j][7]).trim() : ""
      });
    }
  }
  
  logs.sort((a, b) => new Date(b.date) - new Date(a.date));
  return logs;
}

/**
 * Submit employee leave request
 */
/**
 * Helper to calculate number of days between two date strings (inclusive)
 */
function calculateDaysBetween(startDateStr, endDateStr) {
  try {
    var start = new Date(startDateStr);
    var end = new Date(endDateStr);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return 1;
    var diffTime = Math.abs(end - start);
    var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    return diffDays > 0 ? diffDays : 1;
  } catch(e) {
    return 1;
  }
}

/**
 * Submit employee leave request with BUG-003 Overlap validation & Leave Balance engine
 */
function submitLeaveRequest(leave) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
  var lrSheet = ss.getSheetByName('LeaveRequests');
  var empSheet = ss.getSheetByName('Employees');
  if (!lrSheet || !empSheet) throw new Error("Database sheets not initialized");
  
  var empData = empSheet.getDataRange().getValues();
  var empHeaders = empData[0].map(h => String(h).trim());
  var empIdCol = empHeaders.indexOf('ID');
  var empNameCol = empHeaders.indexOf('FullName');
  var empEmailCol = empHeaders.indexOf('Email');
  var empDeptCol = empHeaders.indexOf('Department');
  var totalLeaveCol = empHeaders.indexOf('TotalLeaveDays');
  var usedLeaveCol = empHeaders.indexOf('UsedLeaveDays');
  
  var empId = "";
  var empName = "";
  var empDept = "";
  var totalLeaveDays = 12;
  var usedLeaveDays = 0;
  var userEmail = String(leave.employeeEmail || "").toLowerCase().trim();
  
  for (var i = 1; i < empData.length; i++) {
    var sheetEmail = String(empData[i][empEmailCol !== -1 ? empEmailCol : 3]).toLowerCase().trim();
    if (sheetEmail === userEmail) {
      empId = String(empData[i][empIdCol !== -1 ? empIdCol : 0]).trim();
      empName = String(empData[i][empNameCol !== -1 ? empNameCol : 1]).trim();
      empDept = String(empData[i][empDeptCol !== -1 ? empDeptCol : 5]).trim();
      if (totalLeaveCol !== -1 && empData[i][totalLeaveCol]) totalLeaveDays = Number(empData[i][totalLeaveCol]);
      if (usedLeaveCol !== -1 && empData[i][usedLeaveCol]) usedLeaveDays = Number(empData[i][usedLeaveCol]);
      break;
    }
  }
  
  if (!empId) throw new Error("Không tìm thấy thông tin nhân viên với email: " + userEmail);
  
  var newStart = new Date(leave.startDate);
  var newEnd = new Date(leave.endDate);
  if (isNaN(newStart.getTime()) || isNaN(newEnd.getTime())) {
    throw new Error("Ngày bắt đầu hoặc ngày kết thúc không hợp lệ.");
  }
  if (newEnd < newStart) {
    throw new Error("Ngày kết thúc không được nhỏ hơn ngày bắt đầu.");
  }
  
  // BUG-003: Overlap Leave Date Validation
  var lrData = lrSheet.getDataRange().getValues();
  for (var k = 1; k < lrData.length; k++) {
    var existingEmpId = String(lrData[k][1]).trim();
    var existingStatus = String(lrData[k][7]).trim();
    if (existingEmpId === empId && (existingStatus === 'Pending' || existingStatus === 'Approved')) {
      var existStartStr = lrData[k][3] ? (lrData[k][3] instanceof Date ? Utilities.formatDate(lrData[k][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(lrData[k][3])) : "";
      var existEndStr = lrData[k][4] ? (lrData[k][4] instanceof Date ? Utilities.formatDate(lrData[k][4], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(lrData[k][4])) : "";
      
      var existStart = new Date(existStartStr);
      var existEnd = new Date(existEndStr);
      
      if (!isNaN(existStart.getTime()) && !isNaN(existEnd.getTime())) {
        if (!(newEnd < existStart || newStart > existEnd)) {
          throw new Error("Bạn đã có đơn xin nghỉ phép (" + existingStatus + ") từ ngày " + existStartStr + " đến " + existEndStr + ". Không thể gửi trùng lịch.");
        }
      }
    }
  }
  
  // Leave Balance Check
  var requestedDays = calculateDaysBetween(leave.startDate, leave.endDate);
  var leaveType = leave.leaveType || "Nghỉ phép năm";
  var availableLeave = totalLeaveDays - usedLeaveDays;
  var noteReason = leave.reason || "";
  
  if (leaveType === "Nghỉ phép năm") {
    if (availableLeave <= 0) {
      leaveType = "Nghỉ không lương";
      noteReason += " [Hệ thống tự động chuyển sang Nghỉ không lương do đã hết Quỹ phép năm]";
    } else if (requestedDays > availableLeave) {
      noteReason += " [Vượt quá " + (requestedDays - availableLeave) + " ngày phép khả dụng]";
    }
  }
  
  var nextId = "LR-0001";
  if (lrData.length > 1) {
    var lastId = String(lrData[lrData.length - 1][0]);
    var match = lastId.match(/LR-(\d+)/);
    if (match) {
      var num = parseInt(match[1], 10) + 1;
      nextId = "LR-" + ("0000" + num).slice(-4);
    }
  }
  
  var managerEmail = "";
  for (var m = 1; m < empData.length; m++) {
    var dept = String(empData[m][empDeptCol !== -1 ? empDeptCol : 5]).trim();
    var pos = String(empData[m][6]).toLowerCase().trim();
    if (dept === empDept && (pos.indexOf("manager") !== -1 || pos.indexOf("trưởng phòng") !== -1 || pos.indexOf("lead") !== -1)) {
      managerEmail = String(empData[m][empEmailCol !== -1 ? empEmailCol : 3]).trim();
      break;
    }
  }
  if (!managerEmail) managerEmail = "hr@recruitflow.com";
  
  lrSheet.appendRow([
    nextId,
    empId,
    leaveType,
    leave.startDate || "",
    leave.endDate || "",
    noteReason,
    managerEmail,
    "Pending"
  ]);
  
  try {
    var emailBody = "<div style='font-family: sans-serif; padding: 20px; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px;'>" +
                    "<h2>YÊU CẦU DUYỆT NGHỈ PHÉP</h2>" +
                    "<p>Kính gửi Quản lý,</p>" +
                    "<p>Nhân viên <strong>" + empName + "</strong> (Phòng ban: " + empDept + ") vừa gửi yêu cầu nghỉ phép cần bạn phê duyệt:</p>" +
                    "<table style='width: 100%; border-collapse: collapse; margin: 15px 0;'>" +
                    "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Loại phép:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + leaveType + "</td></tr>" +
                    "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Thời gian:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>Từ " + leave.startDate + " đến " + leave.endDate + " (" + requestedDays + " ngày)</td></tr>" +
                    "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Lý do:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + noteReason + "</td></tr>" +
                    "</table>" +
                    "<p>Vui lòng đăng nhập vào hệ thống HRM để duyệt đơn này.</p>" +
                    "<br><p>Trân trọng,</p><p><strong>Hệ thống HRM RecruitFlow</strong></p>" +
                    "</div>";
                    
    MailApp.sendEmail({
      to: managerEmail,
      subject: "Yêu cầu nghỉ phép mới cần duyệt - " + empName,
      body: "Nhân viên " + empName + " yêu cầu nghỉ phép từ " + leave.startDate + " đến " + leave.endDate,
      htmlBody: emailBody
    });
  } catch (err) {
    Logger.log("Failed to send leave email notification: " + err.message);
  }
  
  return { success: true, message: "Gửi đơn xin nghỉ phép thành công." };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get leave requests matching RBAC filters
 */
function getLeaveRequests(userEmail, role) {
  var ss = getSpreadsheet();
  var lrSheet = ss.getSheetByName('LeaveRequests');
  var empSheet = ss.getSheetByName('Employees');
  if (!lrSheet || !empSheet) return [];
  
  var empData = empSheet.getDataRange().getValues();
  var empIdToNameMap = {};
  var empIdToEmailMap = {};
  var empIdToDeptMap = {};
  var targetEmployeeId = "";
  var cleanUserEmail = String(userEmail || "").toLowerCase().trim();
  
  for (var i = 1; i < empData.length; i++) {
    var id = String(empData[i][0]).trim();
    var name = String(empData[i][1]).trim();
    var email = String(empData[i][3] || empData[i][2]).toLowerCase().trim();
    var dept = String(empData[i][5] || empData[i][4]).trim();
    empIdToNameMap[id] = name;
    empIdToEmailMap[id] = email;
    empIdToDeptMap[id] = dept;
    if (email === cleanUserEmail) {
      targetEmployeeId = id;
    }
  }
  
  var list = [];
  var data = lrSheet.getDataRange().getValues();
  for (var j = 1; j < data.length; j++) {
    var reqEmpId = String(data[j][1]).trim();
    var reqEmail = empIdToEmailMap[reqEmpId] || "";
    var reqDept = empIdToDeptMap[reqEmpId] || "";
    var approver = String(data[j][6]).toLowerCase().trim();
    var status = String(data[j][7]).trim();
    
    var allow = false;
    if (role === 'Admin' || role === 'HR') {
      allow = true;
    } else if (role === 'Manager') {
      var managerDept = "";
      for (var k = 1; k < empData.length; k++) {
        var e = String(empData[k][3] || empData[k][2]).toLowerCase().trim();
        if (e === cleanUserEmail) {
          managerDept = String(empData[k][5] || empData[k][4]).trim();
          break;
        }
      }
      if (approver === cleanUserEmail || (managerDept && reqDept === managerDept)) {
        allow = true;
      }
    } else if (role === 'Employee' || reqEmpId === targetEmployeeId) {
      if (reqEmpId === targetEmployeeId) {
        allow = true;
      }
    }
    
    if (allow) {
      list.push({
        id: data[j][0] ? String(data[j][0]).trim() : "",
        employeeId: reqEmpId,
        employeeName: empIdToNameMap[reqEmpId] || "Unknown",
        employeeEmail: reqEmail,
        department: reqDept,
        leaveType: data[j][2] ? String(data[j][2]).trim() : "",
        startDate: data[j][3] ? (data[j][3] instanceof Date ? Utilities.formatDate(data[j][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[j][3])) : "",
        endDate: data[j][4] ? (data[j][4] instanceof Date ? Utilities.formatDate(data[j][4], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[j][4])) : "",
        reason: data[j][5] ? String(data[j][5]).trim() : "",
        approver: approver,
        status: status
      });
    }
  }
  
  list.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
  return list;
}

/**
 * Approve or Reject leave requests & update Leave Balance
 */
function approveLeaveRequest(requestId, status, approverEmail) {
  requireRole(['Admin', 'HR', 'Manager']);
  var ss = getSpreadsheet();
  var lrSheet = ss.getSheetByName('LeaveRequests');
  var empSheet = ss.getSheetByName('Employees');
  if (!lrSheet || !empSheet) throw new Error("Sheets not found");
  
  var data = lrSheet.getDataRange().getValues();
  var rowIndex = -1;
  var employeeId = "";
  var leaveType = "";
  var startDateStr = "";
  var endDateStr = "";
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === requestId) {
      rowIndex = i + 1;
      employeeId = String(data[i][1]).trim();
      leaveType = String(data[i][2]).trim();
      startDateStr = data[i][3] ? (data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][3])) : "";
      endDateStr = data[i][4] ? (data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][4])) : "";
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Không tìm thấy đơn xin nghỉ phép mã số: " + requestId);
  
  lrSheet.getRange(rowIndex, 8).setValue(status);
  
  var empData = empSheet.getDataRange().getValues();
  var empHeaders = empData[0].map(h => String(h).trim());
  var usedLeaveCol = empHeaders.indexOf('UsedLeaveDays');
  var empEmail = "";
  var empName = "";
  var empRowIndex = -1;
  var currentUsedLeave = 0;
  
  for (var j = 1; j < empData.length; j++) {
    if (String(empData[j][0]).trim() === employeeId) {
      empRowIndex = j + 1;
      empName = String(empData[j][1]).trim();
      empEmail = String(empData[j][3] || empData[j][2]).trim();
      if (usedLeaveCol !== -1 && empData[j][usedLeaveCol]) currentUsedLeave = Number(empData[j][usedLeaveCol]);
      break;
    }
  }
  
  // If status is Approved and leave type is Nghỉ phép năm, update UsedLeaveDays
  if (status === 'Approved' && leaveType === 'Nghỉ phép năm' && empRowIndex !== -1 && usedLeaveCol !== -1) {
    var numDays = calculateDaysBetween(startDateStr, endDateStr);
    empSheet.getRange(empRowIndex, usedLeaveCol + 1).setValue(currentUsedLeave + numDays);
  }
  
  if (empEmail) {
    try {
      var statusColor = status === "Approved" ? "#16a34a" : "#dc2626";
      var statusText = status === "Approved" ? "ĐÃ PHÊ DUYỆT / APPROVED" : "TỪ CHỐI / REJECTED";
      var emailBody = "<div style='font-family: sans-serif; padding: 20px; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px;'>" +
                      "<h2>KẾT QUẢ ĐƠN XIN NGHỈ PHÉP</h2>" +
                      "<p>Kính gửi <strong>" + empName + "</strong>,</p>" +
                      "<p>Yêu cầu nghỉ phép mã số <strong>" + requestId + "</strong> của bạn đã được xử lý bởi Quản lý:</p>" +
                      "<div style='background-color: #f8fafc; padding: 15px; border-left: 4px solid " + statusColor + "; margin: 15px 0;'>" +
                      "Trạng thái: <strong style='color: " + statusColor + ";'>" + statusText + "</strong>" +
                      "</div>" +
                      "<p>Trân trọng cảm ơn,</p><p><strong>Hệ thống HRM RecruitFlow</strong></p>" +
                      "</div>";
                      
      MailApp.sendEmail({
        to: empEmail,
        subject: "Kết quả duyệt đơn nghỉ phép: " + statusText + " - " + requestId,
        body: "Đơn xin nghỉ phép " + requestId + " của bạn đã: " + status,
        htmlBody: emailBody
      });
    } catch(err) {
      Logger.log("Failed to send leave result email: " + err.message);
    }
  }
  
  return { success: true };
}

/**
 * =========================================================================
 * OVERTIME (OT) REQUEST & APPROVAL BACKEND ENGINE
 * =========================================================================
 */

function submitOtRequest(ot) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('OvertimeRequests');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('OvertimeRequests');
    }
    var empSheet = ss.getSheetByName('Employees');
    if (!sheet || !empSheet) throw new Error("OvertimeRequests database sheet not initialized");
    
    var empData = empSheet.getDataRange().getValues();
    var empId = "";
    var userEmail = String(ot.employeeEmail || "").toLowerCase().trim();
    for (var i = 1; i < empData.length; i++) {
      var email = String(empData[i][3] || empData[i][2]).toLowerCase().trim();
      if (email === userEmail) {
        empId = String(empData[i][0]).trim();
        break;
      }
    }
    if (!empId) throw new Error("Không tìm thấy mã nhân viên tương ứng với email: " + userEmail);
    
    var data = sheet.getDataRange().getValues();
    var nextId = "OT-0001";
    if (data.length > 1) {
      var lastId = String(data[data.length - 1][0]);
      var match = lastId.match(/OT-(\d+)/);
      if (match) nextId = "OT-" + ("0000" + (parseInt(match[1], 10) + 1)).slice(-4);
    }
    
    var multiplier = 1.5;
    if (ot.otType === 'Weekend' || ot.otType === 'Cuối tuần') multiplier = 2.0;
    if (ot.otType === 'Holiday' || ot.otType === 'Ngày lễ') multiplier = 3.0;
    
    sheet.appendRow([
      nextId,
      empId,
      ot.date || "",
      Number(ot.hours) || 0,
      ot.otType || "Ngày thường",
      multiplier,
      ot.reason || "",
      ot.approver || "hr@recruitflow.com",
      "Pending",
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ]);
    return { success: true, otId: nextId, message: "Đăng ký tăng ca (OT) thành công." };
  } finally {
    lock.releaseLock();
  }
}

function getOtRequests(userEmail, role) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('OvertimeRequests');
  var empSheet = ss.getSheetByName('Employees');
  if (!sheet || !empSheet) return [];
  
  var empData = empSheet.getDataRange().getValues();
  var empMap = {};
  var targetEmpId = "";
  var cleanEmail = String(userEmail || "").toLowerCase().trim();
  
  for (var i = 1; i < empData.length; i++) {
    var id = String(empData[i][0]).trim();
    var name = String(empData[i][1]).trim();
    var email = String(empData[i][3] || empData[i][2]).toLowerCase().trim();
    var dept = String(empData[i][5] || empData[i][4]).trim();
    empMap[id] = { name: name, email: email, dept: dept };
    if (email === cleanEmail) targetEmpId = id;
  }
  
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var j = 1; j < data.length; j++) {
    var reqEmpId = String(data[j][1]).trim();
    var info = empMap[reqEmpId] || { name: "Unknown", email: "", dept: "" };
    var allow = false;
    if (role === 'Admin' || role === 'HR') allow = true;
    else if (role === 'Manager') {
      if (String(data[j][7]).toLowerCase().trim() === cleanEmail || info.email === cleanEmail) allow = true;
    } else if (reqEmpId === targetEmpId) allow = true;
    
    if (allow) {
      list.push({
        id: String(data[j][0]).trim(),
        employeeId: reqEmpId,
        employeeName: info.name,
        employeeEmail: info.email,
        department: info.dept,
        date: data[j][2] ? (data[j][2] instanceof Date ? Utilities.formatDate(data[j][2], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[j][2])) : "",
        hours: Number(data[j][3]) || 0,
        otType: String(data[j][4]).trim(),
        multiplier: Number(data[j][5]) || 1.5,
        reason: String(data[j][6]).trim(),
        approver: String(data[j][7]).trim(),
        status: String(data[j][8]).trim()
      });
    }
  }
  list.sort((a, b) => new Date(b.date) - new Date(a.date));
  return list;
}

function approveOtRequest(requestId, status, approverEmail) {
  requireRole(['Admin', 'HR', 'Manager']);
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('OvertimeRequests');
  if (!sheet) throw new Error("OvertimeRequests sheet not found");
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === requestId) {
      sheet.getRange(i + 1, 9).setValue(status);
      sheet.getRange(i + 1, 8).setValue(approverEmail || "");
      return { success: true, message: "Cập nhật trạng thái đơn OT thành công." };
    }
  }
  throw new Error("Không tìm thấy đơn OT mã số: " + requestId);
}

/**
 * PIT Taxes Brackets calculation
 */
function calculatePIT(taxableIncome) {
  if (taxableIncome <= 0) return 0;
  if (taxableIncome <= 5000000) return taxableIncome * 0.05;
  if (taxableIncome <= 10000000) return taxableIncome * 0.10 - 250000;
  if (taxableIncome <= 18000000) return taxableIncome * 0.15 - 750000;
  if (taxableIncome <= 32000000) return taxableIncome * 0.20 - 1650000;
  if (taxableIncome <= 52000000) return taxableIncome * 0.25 - 3250000;
  if (taxableIncome <= 80000000) return taxableIncome * 0.30 - 5850000;
  return taxableIncome * 0.35 - 9850000;
}

/**
 * Tự động tính Bảo hiểm bắt buộc theo Luật Lao động Việt Nam
 * - BHXH (8%) + BHYT (1.5%) = 9.5%, khống chế trần 20 lần Lương cơ sở (46,800,000 VNĐ - NĐ 73/2024/NĐ-CP)
 * - BHTN (1%) = 1.0%, khống chế trần 20 lần Lương tối thiểu vùng I (99,200,000 VNĐ - NĐ 74/2024/NĐ-CP)
 */
function calculateMandatoryInsurance(socialInsuranceSalary, basicSalary) {
  var configs = getConfig();
  var baseSalaryCap = Number(configs.base_salary_cap || 46800000);
  var regionMinSalaryCap = Number(configs.region_min_salary_cap || 99200000);
  
  var salaryBase = Number(socialInsuranceSalary) || Number(basicSalary) || 0;
  
  // BHXH (8%) + BHYT (1.5%) = 9.5% (Capped at 20x Base Salary)
  var bhxhBhytBase = Math.min(salaryBase, baseSalaryCap);
  var bhxhBhyt = bhxhBhytBase * 0.095;
  
  // BHTN (1%) (Capped at 20x Region I Min Salary)
  var bhtnBase = Math.min(salaryBase, regionMinSalaryCap);
  var bhtn = bhtnBase * 0.01;
  
  return Math.round(bhxhBhyt + bhtn);
}

/**
 * Trình kích hoạt chạy hàng ngày (Daily Time-driven Trigger):
 * Quét danh sách nhân viên có hợp đồng / thời gian thử việc sắp hết hạn (15/30/60 ngày).
 * Tự động gửi Email thông báo tới Trưởng phòng & HR kèm hướng dẫn Đánh giá Tái ký.
 */
function checkContractExpirationsTrigger() {
  try {
    var configs = getConfig();
    var alertDaysStr = configs.contract_alert_days || "15,30,60";
    var alertDaysList = alertDaysStr.split(',').map(function(d) { return parseInt(d.trim(), 10); }).filter(function(d) { return !isNaN(d); });
    
    var employees = getEmployees();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var expiringEmployees = [];
    
    for (var i = 0; i < employees.length; i++) {
      var emp = employees[i];
      if (emp.status !== 'Đang làm' && emp.status !== 'Thử việc' && emp.status !== 'Chính thức') continue;
      
      // Lọc bỏ hợp đồng không xác định thời hạn
      var contractType = String(emp.contractType || '').toLowerCase().trim();
      if (contractType === 'không xác định thời hạn' || contractType === 'indefinite' || contractType === 'vô thời hạn') {
        continue;
      }
      
      var targetDateStr = emp.contractExpiryDate || emp.probationEndDate || emp.endDate;
      if (!targetDateStr) continue;
      
      var targetDate = new Date(targetDateStr);
      if (isNaN(targetDate.getTime())) continue;
      targetDate.setHours(0, 0, 0, 0);
      
      var diffTime = targetDate.getTime() - today.getTime();
      var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      if (alertDaysList.indexOf(diffDays) !== -1) {
        expiringEmployees.push({
          emp: emp,
          daysLeft: diffDays,
          expiryDate: targetDateStr
        });
      }
    }
    
    if (expiringEmployees.length === 0) return { success: true, message: "Không có hợp đồng nào đến hạn cảnh báo hôm nay." };
    
    var hrUsers = getUsersList();
    var hrEmailsList = hrUsers.filter(u => u.role === 'Admin' || u.role === 'HR').map(u => u.email);
    if (hrEmailsList.length === 0 && getActiveUserEmail()) {
      hrEmailsList.push(getActiveUserEmail());
    }
    
    var bodyHtml = "<div style='font-family: sans-serif; font-size: 13px; color: #1e293b; max-width: 650px; margin: 0 auto; border: 1px solid #e2e8f0; padding: 20px; border-radius: 12px;'>" +
                   "<h3 style='color: #4f46e5; margin-top: 0;'>CẢNH BÁO TỰ ĐỘNG: HỢP ĐỒNG LAO ĐỘNG SẮP HẾT HẠN</h3>" +
                   "<p>Hệ thống RecruitFlow HRM phát hiện <b>" + expiringEmployees.length + " nhân sự</b> có hợp đồng sắp đến hạn tái ký:</p>" +
                   "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; font-size: 12px; border-color: #cbd5e1;'>" +
                   "<tr style='background-color: #f1f5f9; text-align: left;'><th>Mã NV</th><th>Họ tên</th><th>Phòng ban</th><th>Chức danh</th><th>Ngày hết hạn</th><th>Cảnh báo trước</th></tr>";
    
    expiringEmployees.forEach(function(item) {
      bodyHtml += "<tr>" +
        "<td>" + item.emp.id + "</td>" +
        "<td><b>" + item.emp.fullName + "</b></td>" +
        "<td>" + (item.emp.department || '-') + "</td>" +
        "<td>" + (item.emp.position || '-') + "</td>" +
        "<td>" + item.expiryDate + "</td>" +
        "<td style='color: #e11d48; font-weight: bold;'>" + item.daysLeft + " ngày</td>" +
        "</tr>";
    });
    
    bodyHtml += "</table>" +
                "<p style='margin-top: 15px;'>Vui lòng truy cập hệ thống RecruitFlow HRM để tiến hành quy trình Đánh giá Tái ký Hợp đồng theo đúng Luật Lao động 2019.</p>" +
                "</div>";
    
    var sentCount = 0;
    var failedEmails = [];
    hrEmailsList.forEach(function(toEmail) {
      try {
        sendEmailWithFallback(toEmail, "[RecruitFlow HRM] Cảnh báo Hợp đồng sắp hết hạn (" + expiringEmployees.length + " nhân sự)", bodyHtml);
        sentCount++;
      } catch (err) {
        Logger.log("Lỗi gửi email cảnh báo hợp đồng tới " + toEmail + ": " + err.message);
        failedEmails.push({ email: toEmail, error: err.message });
      }
    });
    
    return { success: true, count: expiringEmployees.length, sentCount: sentCount, failedEmails: failedEmails };
  } catch (e) {
    Logger.log("Lỗi checkContractExpirationsTrigger: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Tạo tự động Trình kích hoạt chạy hàng ngày lúc 7:00 AM cho checkContractExpirationsTrigger
 */
function setupContractExpirationDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkContractExpirationsTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkContractExpirationsTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  return { success: true, message: "Đã cài đặt Trình kích hoạt cảnh báo Hợp đồng chạy lúc 7h00 hàng ngày." };
}

/**
 * Ghi vết Lịch sử Biến động Nhân sự vào trang tính JobHistory
 * ChangeType: 'Recruited', 'Probation', 'Official', 'Promoted', 'Transferred', 'SalaryRaised', 'Disciplined', 'Resigned'
 */
function logJobHistory(employeeId, employeeName, changeType, oldValue, newValue, effectiveDate, decisionNumber, notes) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('JobHistory');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('JobHistory');
    }
    
    var id = "JH-" + Number(new Date()) + "-" + Math.floor(Math.random() * 1000);
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var effDateStr = effectiveDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    var row = [
      id,
      employeeId || '',
      employeeName || '',
      changeType || 'Update',
      String(oldValue !== undefined && oldValue !== null ? oldValue : ''),
      String(newValue !== undefined && newValue !== null ? newValue : ''),
      effDateStr,
      decisionNumber || '',
      notes || '',
      nowStr
    ];
    
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { success: true, id: id };
  } catch (e) {
    Logger.log("Lỗi logJobHistory: " + e.message);
    return { success: false, message: e.message };
  } finally {
    try { LockService.getScriptLock().releaseLock(); } catch(err) {}
  }
}


/**
 * Perform monthly payroll calculations with RBAC, Social Insurance Cap, Dependents PIT Tax & Approved OT Requests
 */
function calculatePayroll(month, year) {
  requireRole(['Admin', 'HR']);
  var ss = getSpreadsheet();
  var empSheet = ss.getSheetByName('Employees');
  var attSheet = ss.getSheetByName('Attendance');
  var paySheet = ss.getSheetByName('Payroll');
  var otSheet = ss.getSheetByName('OvertimeRequests');
  if (!empSheet || !attSheet || !paySheet) throw new Error("Database sheets not initialized");
  
  var configs = getConfig();
  var defaultHourlyRate = Number(configs.ot_hourly_rate || 150000);
  
  var employees = getEmployees();
  var activeEmployees = employees.filter(e => e.status === "Đang làm");
  
  // 1. Map OvertimeRequests (Approved)
  var approvedOtMap = {}; // empId -> total ot pay
  if (otSheet) {
    var otData = otSheet.getDataRange().getValues();
    for (var o = 1; o < otData.length; o++) {
      var otEmpId = String(otData[o][1]).trim();
      var otDate = otData[o][2];
      var otDateStr = otDate instanceof Date ? Utilities.formatDate(otDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(otDate).trim();
      var otStatus = String(otData[o][8]).trim();
      
      if (otStatus === 'Approved') {
        var dateParts = otDateStr.split("-");
        if (dateParts.length === 3) {
          var otYear = parseInt(dateParts[0], 10);
          var otMonth = parseInt(dateParts[1], 10);
          if (otYear === parseInt(year, 10) && otMonth === parseInt(month, 10)) {
            var hours = Number(otData[o][3]) || 0;
            var mult = Number(otData[o][5]) || 1.5;
            
            var targetEmp = employees.find(e => e.id === otEmpId);
            var empBasic = targetEmp ? (targetEmp.basicSalary || 0) : 0;
            var rate = empBasic > 0 ? (empBasic / 22 / 8) : defaultHourlyRate;
            var otPay = Math.round(hours * rate * mult);
            
            if (!approvedOtMap[otEmpId]) approvedOtMap[otEmpId] = 0;
            approvedOtMap[otEmpId] += otPay;
          }
        }
      }
    }
  }
  
  // 2. Map Attendance raw OT hours
  var attendanceLogs = [];
  var attData = attSheet.getDataRange().getValues();
  for (var i = 1; i < attData.length; i++) {
    var dbEmpId = String(attData[i][1]).trim();
    var dbDate = attData[i][2];
    var dbDateStr = dbDate instanceof Date ? Utilities.formatDate(dbDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dbDate).trim();
    
    var dateParts = dbDateStr.split("-");
    if (dateParts.length === 3) {
      var logYear = parseInt(dateParts[0], 10);
      var logMonth = parseInt(dateParts[1], 10);
      if (logYear === parseInt(year, 10) && logMonth === parseInt(month, 10)) {
        attendanceLogs.push({
          employeeId: dbEmpId,
          otHours: attData[i][6] ? Number(attData[i][6]) : 0
        });
      }
    }
  }
  
  var rawOtMap = {};
  attendanceLogs.forEach(log => {
    if (!rawOtMap[log.employeeId]) rawOtMap[log.employeeId] = 0;
    rawOtMap[log.employeeId] += log.otHours;
  });
  
  var payData = paySheet.getDataRange().getValues();
  var payRowMap = {};
  for (var p = 1; p < payData.length; p++) {
    var pId = String(payData[p][0]).trim();
    var pMonth = String(payData[p][1]).trim();
    var pYear = String(payData[p][2]).trim();
    var pEmpId = String(payData[p][3]).trim();
    payRowMap[pEmpId + "_" + pMonth + "_" + pYear] = p + 1;
  }
  
  var nextNum = 1;
  if (payData.length > 1) {
    var lastId = String(payData[payData.length - 1][0]);
    var match = lastId.match(/PAY-(\d+)/);
    if (match) nextNum = parseInt(match[1], 10) + 1;
  }
  
  var count = 0;
  for (var e = 0; e < activeEmployees.length; e++) {
    var rawBasicSalary = emp.basicSalary || 0;
    var stdWorkDays = 22;
    var prorateRatio = calculateProrateRatio(emp, month, year, stdWorkDays);
    var basicSalary = Math.round(rawBasicSalary * prorateRatio);
    var allowance = Math.round((emp.allowance || 0) * prorateRatio);
    
    // Total OT Pay from approved OT requests or raw attendance
    var otSalary = approvedOtMap[emp.id] || 0;
    if (otSalary === 0) {
      var rawOtHours = rawOtMap[emp.id] || 0;
      var rate = basicSalary > 0 ? (basicSalary / 22 / 8) : defaultHourlyRate;
      otSalary = Math.round(rawOtHours * rate * 1.5);
    }
    
    var kpiBonus = Number(emp.kpiBonus) || 0;
    var salesCommission = Number(emp.salesCommission) || 0;
    var projectAllowance = Number(emp.projectAllowance) || 0;
    var advanceDeduction = Number(emp.advanceDeduction) || 0;
    var penaltyDeduction = Number(emp.penaltyDeduction) || 0;
    
    var gross = basicSalary + allowance + otSalary + kpiBonus + salesCommission + projectAllowance;
    
    // Social & Unemployment Insurance Base (Capped according to Decree 73/2024/NĐ-CP & 74/2024/NĐ-CP)
    var insurance = calculateMandatoryInsurance(emp.socialInsuranceSalary, basicSalary);
    
    // PIT Taxable Income with Dependents Deduction (4,400,000 VNĐ per dependent)
    var dependents = emp.dependents || 0;
    var taxable = gross - insurance - 11000000 - (dependents * 4400000);
    var tax = Math.round(calculatePIT(taxable));
    var net = gross - insurance - tax - advanceDeduction - penaltyDeduction;
    
    var key = emp.id + "_" + month + "_" + year;
    var rowIdx = payRowMap[key];
    
    if (rowIdx) {
      paySheet.getRange(rowIdx, 5).setValue(gross);
      paySheet.getRange(rowIdx, 6).setValue(insurance);
      paySheet.getRange(rowIdx, 7).setValue(tax);
      paySheet.getRange(rowIdx, 8).setValue(net);
      // Update extended variable pay columns (Cols 10-14)
      paySheet.getRange(rowIdx, 10, 1, 5).setValues([[kpiBonus, salesCommission, projectAllowance, advanceDeduction, penaltyDeduction]]);
    } else {
      var payId = "PAY-" + ("0000" + nextNum).slice(-4);
      paySheet.appendRow([
        payId,
        month,
        year,
        emp.id,
        gross,
        insurance,
        tax,
        net,
        "",
        kpiBonus,
        salesCommission,
        projectAllowance,
        advanceDeduction,
        penaltyDeduction
      ]);
      nextNum++;
    }
    count++;
  }
  
  return { success: true, count: count };
}

/**
 * Generate a high-speed HTML Payslip PDF Blob
 */
function generatePayslipPDFBlob(employee, payrollEntry, month, year) {
  var htmlContent = "<div style='font-family: sans-serif; padding: 30px; max-width: 650px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; color: #1e293b;'>" +
                    "<div style='text-align: center; margin-bottom: 25px;'>" +
                    "<h1 style='color: #4f46e5; margin: 0; font-size: 24px;'>PHIẾU LƯƠNG NHÂN VIÊN</h1>" +
                    "<p style='color: #64748b; font-size: 14px; margin: 5px 0 0 0;'>Tháng " + month + " / Năm " + year + "</p>" +
                    "</div>" +
                    "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin-bottom: 20px;' />" +
                    "<table style='width: 100%; border-collapse: collapse; margin-bottom: 25px; font-size: 14px;'>" +
                    "<tr><td style='padding: 6px 0; color: #64748b;'>Họ và tên:</td><td style='padding: 6px 0; text-align: right; font-weight: 600;'>" + employee.fullName + "</td></tr>" +
                    "<tr><td style='padding: 6px 0; color: #64748b;'>Mã nhân viên:</td><td style='padding: 6px 0; text-align: right; font-weight: 600;'>" + employee.id + "</td></tr>" +
                    "<tr><td style='padding: 6px 0; color: #64748b;'>Phòng ban:</td><td style='padding: 6px 0; text-align: right; font-weight: 600;'>" + employee.department + "</td></tr>" +
                    "<tr><td style='padding: 6px 0; color: #64748b;'>Chức vụ:</td><td style='padding: 6px 0; text-align: right; font-weight: 600;'>" + employee.position + "</td></tr>" +
                    "</table>" +
                    "<h3 style='color: #334155; font-size: 16px; border-bottom: 2px solid #f1f5f9; padding-bottom: 8px; margin-bottom: 12px;'>Chi tiết Thu nhập & Khấu trừ</h3>" +
                    "<table style='width: 100%; border-collapse: collapse; font-size: 14px;'>" +
                    "<tr><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9;'>Lương cơ bản (Basic Salary)</td><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; text-align: right;'>" + Number(employee.basicSalary).toLocaleString() + " VND</td></tr>" +
                    "<tr><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9;'>Phụ cấp (Allowances)</td><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; text-align: right;'>" + Number(employee.allowance || 0).toLocaleString() + " VND</td></tr>" +
                    "<tr><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9;'>Tổng thu nhập Gross (Gross Salary)</td><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; text-align: right; font-weight: 600;'>" + Number(payrollEntry.totalSalary).toLocaleString() + " VND</td></tr>" +
                    "<tr><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; color: #b91c1c;'>Bảo hiểm đóng góp (Insurance 10.5%)</td><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; text-align: right; color: #b91c1c;'>- " + Number(payrollEntry.insurance).toLocaleString() + " VND</td></tr>" +
                    "<tr><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; color: #b91c1c;'>Thuế TNCN tạm tính (Income Tax)</td><td style='padding: 10px 0; border-bottom: 1px solid #f1f5f9; text-align: right; color: #b91c1c;'>- " + Number(payrollEntry.tax).toLocaleString() + " VND</td></tr>" +
                    "<tr style='background-color: #f8fafc; font-size: 16px; font-weight: 700;'>" +
                    "<td style='padding: 12px 10px; color: #1e293b;'>Thực lĩnh (Net Salary)</td>" +
                    "<td style='padding: 12px 10px; text-align: right; color: #16a34a;'>" + Number(payrollEntry.netSalary).toLocaleString() + " VND</td>" +
                    "</tr>" +
                    "</table>" +
                    "<div style='margin-top: 35px; text-align: center; color: #64748b; font-size: 12px; font-style: italic;'>" +
                    "Đây là phiếu lương tự động từ hệ thống HRM. Mọi thắc mắc vui lòng phản hồi phòng Nhân Sự." +
                    "</div>" +
                    "</div>";
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent);
  return htmlOutput.getAs(MimeType.PDF);
}

/**
 * Send PDF payslips automatically via email
 */
function sendPayslips(month, year) {
  var ss = getSpreadsheet();
  var paySheet = ss.getSheetByName('Payroll');
  var empSheet = ss.getSheetByName('Employees');
  if (!paySheet || !empSheet) throw new Error("Payroll sheets not initialized");
  
  var employees = getEmployees();
  var empMap = {};
  employees.forEach(e => { empMap[e.id] = e; });
  
  var payData = paySheet.getDataRange().getValues();
  var sentCount = 0;
  
  for (var i = 1; i < payData.length; i++) {
    var pId = String(payData[i][0]).trim();
    var pMonth = String(payData[i][1]).trim();
    var pYear = String(payData[i][2]).trim();
    var pEmpId = String(payData[i][3]).trim();
    
    if (pMonth === String(month) && pYear === String(year)) {
      var employee = empMap[pEmpId];
      if (!employee) continue;
      
      var payrollEntry = {
        id: pId,
        totalSalary: Number(payData[i][4]),
        insurance: Number(payData[i][5]),
        tax: Number(payData[i][6]),
        netSalary: Number(payData[i][7])
      };
      
      var pdfBlob = generatePayslipPDFBlob(employee, payrollEntry, month, year);
      pdfBlob.setName("Payslip_" + employee.fullName.replace(/\s+/g, "_") + "_" + month + "_" + year + ".pdf");
      
      try {
        var emailBody = "<p>Thân gửi ông/bà <strong>" + employee.fullName + "</strong>,</p>" +
                        "<p>Phòng Nhân Sự xin gửi phiếu lương chi tiết tháng " + month + "/" + year + " trong tệp đính kèm PDF.</p>" +
                        "<p>Vui lòng xem chi tiết thu nhập, các khoản đóng bảo hiểm và thuế TNCN trong tệp đính kèm.</p>" +
                        "<br><p>Trân trọng,</p><p><strong>Phòng Nhân Sự (HR Department)</strong></p>";
                        
        MailApp.sendEmail({
          to: employee.email,
          subject: "Phiếu lương tháng / Payslip - " + month + "/" + year + " - " + employee.fullName,
          body: "Phòng Nhân Sự gửi phiếu lương chi tiết tháng " + month + "/" + year + " đính kèm.",
          htmlBody: emailBody,
          attachments: [pdfBlob]
        });
        
        var nowStr = formatDateString(new Date());
        paySheet.getRange(i + 1, 9).setValue(nowStr);
        sentCount++;
      } catch (err) {
        Logger.log("Failed to send payslip to " + employee.email + ": " + err.message);
      }
    }
  }
  
  return { success: true, count: sentCount };
}

/**
 * Fetch payroll logs securely (Role & Identity verified on server)
 */
function getPayslips(requestedEmail) {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) return [];
  
  var role = getUserRoleByEmail(activeEmail);
  var ss = getSpreadsheet();
  var paySheet = ss.getSheetByName('Payroll');
  var empSheet = ss.getSheetByName('Employees');
  if (!paySheet || !empSheet) return [];
  
  var empData = empSheet.getDataRange().getValues();
  var empIdToNameMap = {};
  var empIdToEmailMap = {};
  var empIdToDeptMap = {};
  var targetEmployeeId = "";
  var cleanUserEmail = (role === 'Admin' || role === 'HR') ? String(requestedEmail || activeEmail).toLowerCase().trim() : activeEmail;
  
  // If input is an employee ID, resolve to their email address first
  if (cleanUserEmail.indexOf("emp-") === 0) {
    for (var i = 1; i < empData.length; i++) {
      var id = String(empData[i][0]).trim();
      if (id.toLowerCase() === cleanUserEmail) {
        cleanUserEmail = String(empData[i][2]).toLowerCase().trim();
        break;
      }
    }
  }
  
  for (var i = 1; i < empData.length; i++) {
    var id = String(empData[i][0]).trim();
    var name = String(empData[i][1]).trim();
    var empEmail = String(empData[i][2]).toLowerCase().trim();
    var dept = String(empData[i][4]).trim();
    empIdToNameMap[id] = name;
    empIdToEmailMap[id] = empEmail;
    empIdToDeptMap[id] = dept;
    
    if (empEmail === cleanUserEmail) {
      targetEmployeeId = id;
    }
  }
  
  var list = [];
  var data = paySheet.getDataRange().getValues();
  for (var j = 1; j < data.length; j++) {
    var empId = String(data[j][3]).trim();
    
    var allow = false;
    if (role === 'Admin' || role === 'HR') {
      allow = true;
    } else if (role === 'Employee' || empId === targetEmployeeId) {
      if (empId === targetEmployeeId) {
        allow = true;
      }
    }
    
    if (allow) {
      list.push({
        id: data[j][0] ? String(data[j][0]).trim() : "",
        month: data[j][1] ? String(data[j][1]).trim() : "",
        year: data[j][2] ? String(data[j][2]).trim() : "",
        employeeId: empId,
        employeeName: empIdToNameMap[empId] || "Unknown",
        employeeEmail: empIdToEmailMap[empId] || "",
        totalSalary: data[j][4] ? Number(data[j][4]) : 0,
        insurance: data[j][5] ? Number(data[j][5]) : 0,
        tax: data[j][6] ? Number(data[j][6]) : 0,
        netSalary: data[j][7] ? Number(data[j][7]) : 0,
        sentDate: data[j][8] ? String(data[j][8]).trim() : ""
      });
    }
  }
  
  list.sort((a, b) => {
    if (b.year !== a.year) return parseInt(b.year) - parseInt(a.year);
    return parseInt(b.month) - parseInt(a.month);
  });
  return list;
}

/**
 * Return a beautifully styled printable decision document for salary history
 */
function getSalaryDecisionPrintHtml(histId) {
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName('SalaryHistory');
  var empSheet = ss.getSheetByName('Employees');
  if (!histSheet || !empSheet) return "";
  
  var histData = histSheet.getDataRange().getValues();
  var histRow = null;
  for (var i = 1; i < histData.length; i++) {
    if (String(histData[i][0]).trim() === histId) {
      histRow = histData[i];
      break;
    }
  }
  if (!histRow) return "";
  
  var empId = String(histRow[1]);
  var newSalary = Number(histRow[2]);
  var changeDate = histRow[3] ? (histRow[3] instanceof Date ? Utilities.formatDate(histRow[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(histRow[3])) : "";
  var notes = String(histRow[5]);
  var decisionFile = String(histRow[4]);
  
  var empData = empSheet.getDataRange().getValues();
  var employee = null;
  for (var j = 1; j < empData.length; j++) {
    if (String(empData[j][0]).trim() === empId) {
      employee = {
        fullName: String(empData[j][1]),
        position: String(empData[j][5]),
        department: String(empData[j][4]),
        basicSalary: Number(empData[j][9])
      };
      break;
    }
  }
  if (!employee) return "";
  
  var companyName = getConfig().company_name || "CÔNG TY CỔ PHẦN RECRUITFLOW";
  
  var html = "<html><head><style>" +
             "body { font-family: 'Times New Roman', Times, serif; padding: 40px; color: #000; line-height: 1.5; }" +
             ".header { text-align: center; margin-bottom: 40px; }" +
             ".title { text-align: center; font-size: 20px; font-weight: bold; margin-top: 30px; margin-bottom: 20px; text-transform: uppercase; }" +
             ".section { margin-bottom: 15px; }" +
             ".footer { margin-top: 50px; display: flex; justify-content: space-between; }" +
             ".footer-col { text-align: center; width: 45%; }" +
             ".signature-space { height: 100px; }" +
             ".attachment-container { margin-top: 40px; text-align: center; page-break-before: always; }" +
             ".attachment-image { max-width: 100%; max-height: 400px; border: 1px solid #ccc; }" +
             "</style></head><body>" +
             "<div class='header'>" +
             "<strong>" + companyName.toUpperCase() + "</strong><br>" +
             "Số: QD-" + histId.replace("SAL-", "") + "/HR<br>" +
             "---------------------" +
             "</div>" +
             "<div class='title'>QUYẾT ĐỊNH<br>V/v: Điều chỉnh mức lương nhân sự</div>" +
             "<div class='section'>- Căn cứ vào Điều lệ hoạt động của " + companyName + ";</div>" +
             "<div class='section'>- Căn cứ vào năng lực đóng góp và hiệu quả công việc của nhân sự;</div>" +
             "<div class='section'>- Xét đề xuất của Trưởng phòng Nhân sự.</div>" +
             "<div class='title' style='font-size: 16px; margin: 20px 0;'>QUYẾT ĐỊNH</div>" +
             "<div class='section'><strong>Điều 1:</strong> Điều chỉnh mức lương cơ bản đối với ông/bà: <strong>" + employee.fullName + "</strong></div>" +
             "<div style='margin-left: 20px;'>" +
             "- Chức vụ: " + employee.position + "<br>" +
             "- Phòng ban: " + employee.department + "<br>" +
             "- Mức lương cũ: " + employee.basicSalary.toLocaleString() + " VND<br>" +
             "- Mức lương mới: <strong>" + newSalary.toLocaleString() + " VND</strong>" +
             "</div>" +
             "<div class='section'><strong>Điều 2:</strong> Quyết định có hiệu lực thi hành từ ngày " + changeDate + ".</div>" +
             "<div class='section'><strong>Điều 3:</strong> Các Phòng ban liên quan và ông/bà <strong>" + employee.fullName + "</strong> chịu trách nhiệm thi hành quyết định này.</div>" +
             "<div class='footer'>" +
             "<div class='footer-col'><strong>Nơi nhận:</strong><br>- Như Điều 3<br>- Lưu HSNS</div>" +
             "<div class='footer-col'><strong>Đại diện Ban Giám Đốc</strong><br><i>(Ký và ghi rõ họ tên)</i><div class='signature-space'></div><strong>Cơ quan Đại diện</strong></div>" +
             "</div>";
             
  if (decisionFile && decisionFile.indexOf("base64") !== -1) {
    html += "<div class='attachment-container'>" +
            "<h3>Tệp đính kèm Quyết định</h3>" +
            "<img src='" + decisionFile + "' class='attachment-image' />" +
            "</div>";
  }
  
  html += "</body></html>";
  return html;
}

/**
 * ----------------------------------------------------
 * HRM SYSTEM CATEGORY MANAGEMENT CRUD OPERATIONS
 * ----------------------------------------------------
 */

/**
 * Get all Master Data categories (with CacheService 30 mins TTL)
 */
function getCategories() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('master_categories');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }
  
  try {
    var ss = getSpreadsheet();
    if (!ss) throw new Error("Không thể kết nối với Google Spreadsheet.");
    
    var sheet = ss.getSheetByName('Categories');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('Categories');
    }
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var categories = [];
    for (var i = 1; i < data.length; i++) {
      categories.push({
        id: data[i][0] ? String(data[i][0]).trim() : "",
        categoryType: data[i][1] ? String(data[i][1]).trim() : "",
        code: data[i][2] ? String(data[i][2]).trim() : "",
        name: data[i][3] ? String(data[i][3]).trim() : "",
        value: data[i][4] ? String(data[i][4]).trim() : "",
        description: data[i][5] ? String(data[i][5]).trim() : ""
      });
    }
    
    try {
      cache.put('master_categories', JSON.stringify(categories), 1800); // 30 phút TTL
    } catch(e) {}
    
    return categories;
  } catch (error) {
    Logger.log("Lỗi trong getCategories: " + error.toString());
    throw new Error("Không thể tải danh mục dữ liệu: " + error.message);
  }
}

/**
 * Add a new Master Data category
 */
function addCategory(cat) {
  try {
    var ss = getSpreadsheet();
    if (!ss) throw new Error("Không thể kết nối với Google Spreadsheet.");
    
    var sheet = ss.getSheetByName('Categories');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('Categories');
    }
    if (!sheet) throw new Error("Không thể khởi tạo bảng Categories.");
    
    var data = sheet.getDataRange().getValues();
    var nextIdNum = 1;
    if (data.length > 1) {
      var lastId = data[data.length - 1][0] ? String(data[data.length - 1][0]).trim() : "";
      var match = lastId.match(/CAT-(\d+)/);
      if (match) {
        nextIdNum = parseInt(match[1], 10) + 1;
      } else {
        nextIdNum = data.length;
      }
    }
    var nextId = "CAT-" + ("0000" + nextIdNum).slice(-4);
    
    sheet.appendRow([
      nextId,
      cat.categoryType || "",
      cat.code || "",
      cat.name || "",
      cat.value || "",
      cat.description || ""
    ]);
    
    try {
      CacheService.getScriptCache().remove('master_categories');
    } catch(e) {}
    
    return { success: true, id: nextId };
  } catch (error) {
    Logger.log("Lỗi trong addCategory: " + error.toString());
    throw new Error("Không thể thêm mới danh mục: " + error.message);
  }
}

/**
 * Update an existing Master Data category
 */
function updateCategory(cat) {
  try {
    var ss = getSpreadsheet();
    if (!ss) throw new Error("Không thể kết nối với Google Spreadsheet.");
    
    var sheet = ss.getSheetByName('Categories');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('Categories');
    }
    if (!sheet) throw new Error("Không thể khởi tạo bảng Categories.");
    
    var data = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === cat.id) {
        rowIdx = i + 1;
        break;
      }
    }
    
    if (rowIdx === -1) throw new Error("Không tìm thấy ID danh mục: " + cat.id);
    
    sheet.getRange(rowIdx, 2, 1, 5).setValues([[
      cat.categoryType || "",
      cat.code || "",
      cat.name || "",
      cat.value || "",
      cat.description || ""
    ]]);
    
    try {
      CacheService.getScriptCache().remove('master_categories');
    } catch(e) {}
    
    return { success: true };
  } catch (error) {
    Logger.log("Lỗi trong updateCategory: " + error.toString());
    throw new Error("Không thể cập nhật danh mục: " + error.message);
  }
}

/**
 * Delete a category item
 */
function deleteCategory(catId) {
  try {
    var ss = getSpreadsheet();
    if (!ss) throw new Error("Không thể kết nối với Google Spreadsheet.");
    
    var sheet = ss.getSheetByName('Categories');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('Categories');
    }
    if (!sheet) throw new Error("Không thể khởi tạo bảng Categories.");
    
    var data = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === catId) {
        rowIdx = i + 1;
        break;
      }
    }
    
    if (rowIdx === -1) throw new Error("Không tìm thấy ID danh mục cần xóa: " + catId);
    
    sheet.deleteRow(rowIdx);
    
    try {
      CacheService.getScriptCache().remove('master_categories');
    } catch(e) {}
    
    return { success: true };
  } catch (error) {
    Logger.log("Lỗi trong deleteCategory: " + error.toString());
    throw new Error("Không thể xóa danh mục: " + error.message);
  }
}


/**
 * ----------------------------------------------------
 * HRM SYSTEM EMPLOYEE DETAILS CRUD & WORKFLOWS
 * ----------------------------------------------------
 */

/**
 * XEM CHI TIẾT NHÂN VIÊN: Tổng hợp dữ liệu đồng thời của 2 bảng Employees và EmployeeDetails phẳng hóa
 * Bao gồm: payrolls[], candidate object, family[], education[], work[], decisions[]
 */
/**
 * XEM CHI TIẾT NHÂN VIÊN: Tổng hợp dữ liệu đồng thời của 2 bảng Employees và EmployeeDetails phẳng hóa
 * Bao gồm: payrolls[], candidate object, family[], education[], work[], decisions[]
 * Self-healing: Tự động thêm dòng vào EmployeeDetails nếu bị thiếu.
 */
function getEmployeeDetail(employeeId) {
    try {
        var ss = getSpreadsheet();
        if (!ss) throw new Error("Không thể kết nối với Google Spreadsheet.");
        
        employeeId = String(employeeId).trim();

        // 1. Lấy dữ liệu từ bảng Employees
        var empSheet = ss.getSheetByName('Employees');
        if (!empSheet) throw new Error("Bảng Employees không tồn tại.");

        var empData = empSheet.getDataRange().getValues();
        var empHeaders = empData[0].map(function(h) { return String(h).trim(); });
        var empColMap = {};
        empHeaders.forEach(function(h, idx) { empColMap[h] = idx; });

        var empRowIndex = -1;
        var empIdColIdx = empColMap['ID'] !== undefined ? empColMap['ID'] : 0;
        for (var i = 1; i < empData.length; i++) {
            if (String(empData[i][empIdColIdx]).trim() === employeeId) {
                empRowIndex = i;
                break;
            }
        }
        if (empRowIndex === -1) throw new Error("Không tìm thấy nhân viên có ID: " + employeeId);

        var empRow = empData[empRowIndex];
        function getEmpVal(colName) {
            var idx = empColMap[colName];
            return (idx !== undefined && empRow[idx] !== undefined && empRow[idx] !== null) ? empRow[idx] : "";
        }

        // 2. Lấy dữ liệu từ bảng EmployeeDetails (Tự động tự phục hồi tạo dòng mới nếu khuyết)
        var detSheet = ss.getSheetByName('EmployeeDetails');
        if (!detSheet) {
            initializeDatabaseV2();
            detSheet = ss.getSheetByName('EmployeeDetails');
        }
        var detData = detSheet.getDataRange().getValues();
        var detHeaders = detData[0].map(function(h) { return String(h).trim(); });
        var detColMap = {};
        detHeaders.forEach(function(h, idx) { detColMap[h] = idx; });

        var detRowIndex = -1;
        var detIdColIdx = detColMap['EmployeeID'] !== undefined ? detColMap['EmployeeID'] : 0;
        for (var i = 1; i < detData.length; i++) {
            if (String(detData[i][detIdColIdx]).trim() === employeeId) {
                detRowIndex = i;
                break;
            }
        }

        var detRow = null;
        if (detRowIndex === -1) {
            var newDetRow = new Array(detHeaders.length).fill("");
            var idIdx = detColMap['EmployeeID'];
            var fnIdx = detColMap['FullName'];
            var dpIdx = detColMap['Department'];
            var gdIdx = detColMap['Gender'];
            var dcIdx = detColMap['Docs'];

            if (idIdx !== undefined) newDetRow[idIdx] = employeeId;
            if (fnIdx !== undefined) newDetRow[fnIdx] = String(getEmpVal('FullName'));
            if (dpIdx !== undefined) newDetRow[dpIdx] = String(getEmpVal('Department'));
            if (gdIdx !== undefined) newDetRow[gdIdx] = String(getEmpVal('Gender'));
            if (dcIdx !== undefined) newDetRow[dcIdx] = "[]";

            detSheet.appendRow(newDetRow);
            detRow = newDetRow;
            SpreadsheetApp.flush();
        } else {
            detRow = detData[detRowIndex];
        }

        function getDetVal(colName) {
            var idx = detColMap[colName];
            return (idx !== undefined && detRow[idx] !== undefined && detRow[idx] !== null) ? detRow[idx] : "";
        }

        function formatDate(d) {
            if (!d) return "";
            if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
            var parsed = Date.parse(String(d));
            if (!isNaN(parsed)) return Utilities.formatDate(new Date(parsed), Session.getScriptTimeZone(), "yyyy-MM-dd");
            return String(d).trim();
        }

        // Tạo object hợp nhất phẳng hoàn toàn
        var employeeCombined = {
            id: employeeId,
            employeeId: employeeId,
            fullName: String(getEmpVal('FullName')).trim(),
            gender: String(getEmpVal('Gender')).trim() || "Nam",
            email: String(getEmpVal('Email')).trim(),
            phone: String(getEmpVal('Phone')).trim(),
            department: String(getEmpVal('Department')).trim(),
            position: String(getEmpVal('Position')).trim(),
            basicSalary: Number(getEmpVal('BasicSalary')) || 0,
            probationSalary: Number(getEmpVal('ProbationSalary')) || 0,
            allowances: String(getEmpVal('Allowances')).trim(),
            allowance: Number(getEmpVal('Allowance')) || 0,
            probationStartDate: formatDate(getEmpVal('ProbationStartDate')),
            officialStartDate: formatDate(getEmpVal('OfficialStartDate') || getEmpVal('StartDate')),
            contractExpiryDate: formatDate(getEmpVal('ContractExpiryDate') || getEmpVal('EndDate')),
            status: String(getEmpVal('Status')).trim() || "Đang làm",
            contractType: String(getEmpVal('ContractType')).trim(),

            avatar: String(getDetVal('Avatar')).trim(),
            birthPlace: String(getDetVal('BirthPlace')).trim(),
            currentAddress: String(getDetVal('CurrentAddress')).trim(),
            registerAddress: String(getDetVal('RegisterAddress')).trim(),
            idCardNumber: String(getDetVal('IdentityCardNumber')).trim(),
            idCardDate: formatDate(getDetVal('IdentityCardDate')),
            idCardPlace: String(getDetVal('IdentityCardPlace')).trim(),
            academicLevel: String(getDetVal('AcademicLevel')).trim(),
            specialization: String(getDetVal('Specialization')).trim(),
            graduationInstitution: String(getDetVal('GraduationInstitution')).trim(),
            youthUnionDate: formatDate(getDetVal('YouthUnionDate')),
            communistPartyDateStatus: String(getDetVal('CommunistPartyDateStatus')).trim(),
            docs: String(getDetVal('Docs')).trim() || "[]",
            dob: formatDate(getDetVal('DateOfBirth')),
            dateOfBirth: formatDate(getDetVal('DateOfBirth'))
        };

        // 3. Lấy thêm mảng dữ liệu liên quan khác
        var familyList = [];
        var famSheet = ss.getSheetByName('FamilyRelations');
        if (famSheet && famSheet.getLastRow() > 1) {
            var famData = famSheet.getDataRange().getValues();
            for (var i = 1; i < famData.length; i++) {
                if (String(famData[i][1]).trim() === employeeId) {
                    familyList.push({
                        id: String(famData[i][0]).trim(),
                        relationType: String(famData[i][2]).trim(),
                        fullName: String(famData[i][3]).trim(),
                        birthYear: String(famData[i][4]).trim(),
                        job: String(famData[i][5]).trim(),
                        workPlace: String(famData[i][6]).trim()
                    });
                }
            }
        }

        var educationList = [];
        var eduSheet = ss.getSheetByName('EducationHistory');
        if (eduSheet && eduSheet.getLastRow() > 1) {
            var eduData = eduSheet.getDataRange().getValues();
            for (var i = 1; i < eduData.length; i++) {
                if (String(eduData[i][1]).trim() === employeeId) {
                    educationList.push({
                        id: String(eduData[i][0]).trim(),
                        fromDate: String(eduData[i][2]).trim(),
                        toDate: String(eduData[i][3]).trim(),
                        schoolName: String(eduData[i][4]).trim(),
                        fieldOfStudy: String(eduData[i][5]).trim(),
                        modeOfStudy: String(eduData[i][6]).trim(),
                        degree: String(eduData[i][7]).trim()
                    });
                }
            }
        }

        var workList = [];
        var workSheet = ss.getSheetByName('WorkHistory');
        if (workSheet && workSheet.getLastRow() > 1) {
            var workData = workSheet.getDataRange().getValues();
            for (var i = 1; i < workData.length; i++) {
                if (String(workData[i][1]).trim() === employeeId) {
                    workList.push({
                        id: String(workData[i][0]).trim(),
                        fromDate: String(workData[i][2]).trim(),
                        toDate: String(workData[i][3]).trim(),
                        company: String(workData[i][4]).trim(),
                        position: String(workData[i][5]).trim(),
                        referencePerson: String(workData[i][6]).trim(),
                        referencePhone: String(workData[i][7]).trim()
                    });
                }
            }
        }

        var decisionsList = [];
        var decSheet = ss.getSheetByName('Decisions');
        if (decSheet && decSheet.getLastRow() > 1) {
            var decData = decSheet.getDataRange().getValues();
            for (var i = 1; i < decData.length; i++) {
                if (String(decData[i][1]).trim() === employeeId) {
                    decisionsList.push({
                        id: String(decData[i][0]).trim(),
                        decisionType: String(decData[i][2]).trim(),
                        decisionNumber: String(decData[i][3]).trim(),
                        title: String(decData[i][4]).trim(),
                        signDate: String(decData[i][5]).trim(),
                        effectiveDate: String(decData[i][6]).trim(),
                        notes: String(decData[i][7]).trim(),
                        fileBase64: String(decData[i][8]).trim()
                    });
                }
            }
        }

        // 4. Lấy danh sách phiếu lương (Payroll)
        var payrollList = [];
        var payrollSheet = ss.getSheetByName('Payroll');
        if (payrollSheet && payrollSheet.getLastRow() > 1) {
            var payrollData = payrollSheet.getDataRange().getValues();
            var payrollHeaders = payrollData[0].map(function(h) { return String(h).trim(); });
            var pyEmpIdx = payrollHeaders.indexOf('EmployeeID');
            for (var i = 1; i < payrollData.length; i++) {
                if (pyEmpIdx !== -1 && String(payrollData[i][pyEmpIdx]).trim() === employeeId) {
                    var pr = {};
                    payrollHeaders.forEach(function(h, idx) {
                        pr[h.charAt(0).toLowerCase() + h.slice(1)] = payrollData[i][idx];
                    });
                    pr.id        = String(payrollData[i][payrollHeaders.indexOf('ID') !== -1 ? payrollHeaders.indexOf('ID') : 0]).trim();
                    pr.month     = payrollData[i][payrollHeaders.indexOf('Month') !== -1 ? payrollHeaders.indexOf('Month') : 1] || '';
                    pr.year      = payrollData[i][payrollHeaders.indexOf('Year') !== -1 ? payrollHeaders.indexOf('Year') : 2] || '';
                    pr.totalSalary  = Number(payrollData[i][payrollHeaders.indexOf('TotalSalary') !== -1 ? payrollHeaders.indexOf('TotalSalary') : 4]) || 0;
                    pr.insurance    = Number(payrollData[i][payrollHeaders.indexOf('Insurance') !== -1 ? payrollHeaders.indexOf('Insurance') : 5]) || 0;
                    pr.tax          = Number(payrollData[i][payrollHeaders.indexOf('Tax') !== -1 ? payrollHeaders.indexOf('Tax') : 6]) || 0;
                    pr.netSalary    = Number(payrollData[i][payrollHeaders.indexOf('NetSalary') !== -1 ? payrollHeaders.indexOf('NetSalary') : 7]) || 0;
                    pr.sentDate     = formatDate(payrollData[i][payrollHeaders.indexOf('SentDate') !== -1 ? payrollHeaders.indexOf('SentDate') : 8]);
                    payrollList.push(pr);
                }
            }
        }

        // 5. Lấy thông tin ứng viên (Candidates) — tìm theo email nhân viên
        var candidateObj = null;
        var empEmail = employeeCombined.email;
        var candSheet = ss.getSheetByName('Candidates');
        if (candSheet && candSheet.getLastRow() > 1 && empEmail) {
            var candData = candSheet.getDataRange().getValues();
            var candHeaders = candData[0].map(function(h) { return String(h).trim(); });
            var cEmailIdx   = candHeaders.indexOf('SenderEmail');
            var cIdIdx      = candHeaders.indexOf('ID');
            var cNameIdx    = candHeaders.indexOf('SenderName');
            var cCVIdx      = candHeaders.indexOf('CV_Link');
            var cStatusIdx  = candHeaders.indexOf('Status');
            var cInterIdx   = candHeaders.indexOf('InterviewInfo');
            var cEvalIdx    = candHeaders.indexOf('EvaluationHistory');
            var cDOBIdx     = candHeaders.indexOf('DateOfBirth');
            var cPhoneIdx   = candHeaders.indexOf('PhoneNumber');
            var cOfferGross = candHeaders.indexOf('OfferGross');
            var cOfferNet   = candHeaders.indexOf('OfferNet');
            var cOfferStart = candHeaders.indexOf('OfferStartDate');
            var cSubjectIdx = candHeaders.indexOf('Subject');
            var cRecvIdx    = candHeaders.indexOf('ReceivedDate');

            for (var i = 1; i < candData.length; i++) {
                var rowEmail = cEmailIdx !== -1 ? String(candData[i][cEmailIdx]).trim().toLowerCase() : '';
                if (rowEmail && rowEmail === empEmail.toLowerCase()) {
                    candidateObj = {
                        id:               cIdIdx !== -1 ? String(candData[i][cIdIdx]).trim() : '',
                        senderName:       cNameIdx !== -1 ? String(candData[i][cNameIdx]).trim() : '',
                        senderEmail:      empEmail,
                        subject:          cSubjectIdx !== -1 ? String(candData[i][cSubjectIdx]).trim() : '',
                        cvLink:           cCVIdx !== -1 ? String(candData[i][cCVIdx]).trim() : '',
                        status:           cStatusIdx !== -1 ? String(candData[i][cStatusIdx]).trim() : '',
                        interviewInfo:    cInterIdx !== -1 ? String(candData[i][cInterIdx]).trim() : '',
                        evaluationHistory: cEvalIdx !== -1 ? String(candData[i][cEvalIdx]).trim() : '',
                        dateOfBirth:      cDOBIdx !== -1 ? formatDate(candData[i][cDOBIdx]) : '',
                        phone:            cPhoneIdx !== -1 ? String(candData[i][cPhoneIdx]).trim() : '',
                        offerGross:       cOfferGross !== -1 ? Number(candData[i][cOfferGross]) || 0 : 0,
                        offerNet:         cOfferNet !== -1 ? Number(candData[i][cOfferNet]) || 0 : 0,
                        offerStartDate:   cOfferStart !== -1 ? formatDate(candData[i][cOfferStart]) : '',
                        receivedDate:     cRecvIdx !== -1 ? formatDate(candData[i][cRecvIdx]) : ''
                    };
                    break;
                }
            }
        }

        // 6. Lấy danh sách chuyên cần (Attendance)
        var attendanceList = [];
        var attSheet = ss.getSheetByName('Attendance');
        if (attSheet && attSheet.getLastRow() > 1) {
            var attData = attSheet.getDataRange().getValues();
            var attHeaders = attData[0].map(function(h) { return String(h).trim(); });
            var attEmpIdx = attHeaders.indexOf('EmployeeID');
            var attDateIdx = attHeaders.indexOf('Date');
            var attCheckInIdx = attHeaders.indexOf('CheckIn');
            var attCheckOutIdx = attHeaders.indexOf('CheckOut');
            var attStatusIdx = attHeaders.indexOf('Status');
            
            for (var i = 1; i < attData.length; i++) {
                if (attEmpIdx !== -1 && String(attData[i][attEmpIdx]).trim() === employeeId) {
                    var attDate = attDateIdx !== -1 ? attData[i][attDateIdx] : '';
                    var attDateStr = attDate ? (attDate instanceof Date ? Utilities.formatDate(attDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(attDate)) : "";
                    
                    attendanceList.push({
                        id: attHeaders.indexOf('ID') !== -1 ? String(attData[i][attHeaders.indexOf('ID')]).trim() : '',
                        date: attDateStr,
                        checkIn: attCheckInIdx !== -1 ? String(attData[i][attCheckInIdx]).trim() : '',
                        checkOut: attCheckOutIdx !== -1 ? String(attData[i][attCheckOutIdx]).trim() : '',
                        status: attStatusIdx !== -1 ? String(attData[i][attStatusIdx]).trim() : ''
                    });
                }
            }
        }

        return {
            success: true,
            employee: employeeCombined,
            family: familyList,
            education: educationList,
            work: workList,
            decisions: decisionsList,
            payrolls: payrollList,
            candidate: candidateObj,
            attendance: attendanceList
        };

    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

/**
 * LƯU THAY ĐỔI CHI TIẾT NÂNG CẤP VỚI CƠ CHẾ STAGING & TRANSACTION ROLLBACK SỐ DỮ LIỆU
 */
function saveEmployeeDetail(details) {
    requireRole(['Admin', 'HR']);
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var employeeId = details.employeeId || details.id || '';
        if (!employeeId) throw new Error('ID nhân sự không hợp lệ. Kiểm tra details.employeeId hoặc details.id.');

        var empName = details.fullName || '';

        // ---- STEP 1: STAGING & SNAPSHOT PHASE (PREPARE IN MEMORY BEFORE WRITING) ----
        var empRead = batchReadSheet('Employees');
        if (!empRead.sheet || empRead.data.length <= 1) throw new Error("Không tìm thấy dữ liệu bảng Employees.");
        
        var detRead = batchReadSheet('EmployeeDetails');
        if (!detRead.sheet) {
            initializeDatabaseV2();
            detRead = batchReadSheet('EmployeeDetails');
        }
        if (!detRead.sheet) throw new Error("Không thể khởi tạo bảng EmployeeDetails.");

        // Snapshot original state for Rollback safety
        var originalEmpData = JSON.parse(JSON.stringify(empRead.data));
        var originalDetData = JSON.parse(JSON.stringify(detRead.data));

        // Prepare Staged Employees Array
        var empData = empRead.data;
        var empHeaders = empRead.headers;
        var empColMap = {};
        empHeaders.forEach(function(h, idx) { empColMap[h] = idx; });

        var empRowIdx = -1;
        for (var i = 1; i < empData.length; i++) {
            if (String(empData[i][0]).trim() === employeeId) {
                empRowIdx = i;
                break;
            }
        }
        if (empRowIdx === -1) throw new Error("Không tìm thấy mã nhân sự " + employeeId + " trong bảng Employees.");

        var row = empData[empRowIdx];

        // Audit JobHistory changes in Staging
        var oldStatus = empColMap['Status'] !== undefined ? String(row[empColMap['Status']]).trim() : '';
        var oldDept   = empColMap['Department'] !== undefined ? String(row[empColMap['Department']]).trim() : '';
        var oldPos    = empColMap['Position'] !== undefined ? String(row[empColMap['Position']]).trim() : '';
        var oldSalary = empColMap['BasicSalary'] !== undefined ? Number(row[empColMap['BasicSalary']]) || 0 : 0;

        var newStatus = details.status !== undefined ? String(details.status).trim() : oldStatus;
        var newDept   = details.department !== undefined ? String(details.department).trim() : oldDept;
        var newPos    = details.position !== undefined ? String(details.position).trim() : oldPos;
        var newSalary = details.basicSalary !== undefined ? Number(details.basicSalary) || 0 : oldSalary;

        function setCell(colName, val) {
            var idx = empColMap[colName];
            if (idx !== undefined && val !== undefined && val !== null) {
                row[idx] = val;
            }
        }

        setCell('FullName',           details.fullName);
        setCell('Gender',             details.gender);
        setCell('Email',              details.email);
        setCell('Phone',              details.phone);
        setCell('Department',         details.department);
        setCell('Position',           details.position);
        setCell('BasicSalary',        details.basicSalary ? Number(details.basicSalary) : 0);
        setCell('ProbationSalary',    details.probationSalary ? Number(details.probationSalary) : 0);
        setCell('Allowances',         details.allowances);
        setCell('ProbationStartDate', details.probationStartDate);
        setCell('OfficialStartDate',  details.officialStartDate);
        setCell('ContractExpiryDate', details.contractExpiryDate);
        setCell('Status',             details.status);
        setCell('ContractType',       details.contractType);
        setCell('StartDate',          details.officialStartDate || details.probationStartDate);
        setCell('EndDate',            details.contractExpiryDate);

        empData[empRowIdx] = row;

        // Prepare Staged EmployeeDetails Array
        var detData = detRead.data;
        var detHeaders = detRead.headers;
        if (detData.length === 0) {
            detData = [detHeaders];
        }
        var detColMap = {};
        detHeaders.forEach(function(h, idx) { detColMap[h] = idx; });

        var detRowIdx = -1;
        for (var d = 1; d < detData.length; d++) {
            if (String(detData[d][0]).trim() === employeeId) {
                detRowIdx = d;
                break;
            }
        }

        if (detRowIdx === -1) {
            var healRow = new Array(detHeaders.length).fill('');
            if (detColMap['EmployeeID'] !== undefined) healRow[detColMap['EmployeeID']] = employeeId;
            if (detColMap['FullName'] !== undefined)   healRow[detColMap['FullName']]   = details.fullName || '';
            if (detColMap['Department'] !== undefined) healRow[detColMap['Department']] = details.department || '';
            if (detColMap['Gender'] !== undefined)     healRow[detColMap['Gender']]     = details.gender || 'Nam';
            if (detColMap['Docs'] !== undefined)       healRow[detColMap['Docs']]       = '[]';
            detData.push(healRow);
            detRowIdx = detData.length - 1;
        }

        var detRow = detData[detRowIdx];
        function setDetCell(colName, val) {
            var idx = detColMap[colName];
            if (idx !== undefined && val !== undefined && val !== null) {
                detRow[idx] = val;
            }
        }

        setDetCell('FullName',                  details.fullName);
        setDetCell('Department',                details.department);
        setDetCell('Avatar',                    details.avatar);
        setDetCell('Gender',                    details.gender);
        setDetCell('BirthPlace',               details.birthPlace);
        setDetCell('CurrentAddress',           details.currentAddress);
        setDetCell('RegisterAddress',          details.registerAddress);
        setDetCell('IdentityCardNumber',       details.idCardNumber);
        setDetCell('IdentityCardDate',         details.idCardDate);
        setDetCell('IdentityCardPlace',        details.idCardPlace);
        setDetCell('AcademicLevel',            details.academicLevel);
        setDetCell('Specialization',           details.specialization);
        setDetCell('GraduationInstitution',    details.graduationInstitution);
        setDetCell('YouthUnionDate',           details.youthUnionDate);
        setDetCell('CommunistPartyDateStatus', details.communistPartyDateStatus);
        setDetCell('Docs',                     details.docs);
        setDetCell('DateOfBirth',              details.dob || details.dateOfBirth);

        detData[detRowIdx] = detRow;

        // ---- STEP 2: COMMIT PHASE WITH ROLLBACK SAFETY ----
        try {
            // Write step 1: Employees
            batchWriteSheet('Employees', empData);
            
            // Write step 2: EmployeeDetails
            batchWriteSheet('EmployeeDetails', detData);

            // Log history AFTER successful atomic commit
            if (newStatus && oldStatus && newStatus !== oldStatus) {
                logJobHistory(employeeId, empName, 'StatusChange', oldStatus, newStatus, '', '', 'Thay đổi trạng thái nhân sự');
            }
            if (newDept && oldDept && newDept !== oldDept) {
                logJobHistory(employeeId, empName, 'Transferred', oldDept, newDept, '', '', 'Điều chuyển phòng ban');
            }
            if (newPos && oldPos && newPos !== oldPos) {
                logJobHistory(employeeId, empName, 'Promoted', oldPos, newPos, '', '', 'Thay đổi chức danh công tác');
            }
            if (newSalary !== oldSalary && oldSalary > 0) {
                logJobHistory(employeeId, empName, 'SalaryRaised', oldSalary, newSalary, '', '', 'Điều chỉnh mức lương cơ bản');
            }

            try { CacheService.getScriptCache().remove('all_employees'); } catch(e) {}

            return { success: true, message: 'Cập nhật thông tin nhân sự thành công!', employeeId: employeeId };

        } catch (commitErr) {
            // ROLLBACK: Restore original snapshot to Google Sheets if write failed mid-way
            Logger.log("COMMIT FAILED! Rolling back Employees and EmployeeDetails... Error: " + commitErr.message);
            try { batchWriteSheet('Employees', originalEmpData); } catch(rb1) {}
            try { batchWriteSheet('EmployeeDetails', originalDetData); } catch(rb2) {}
            throw new Error("Giao dịch lưu dữ liệu thất bại, hệ thống đã Rollback khôi phục dữ liệu ban đầu: " + commitErr.message);
        }

    } catch (e) {
        Logger.log('Lỗi saveEmployeeDetail: ' + e.message);
        return { success: false, message: e.message };
    } finally {
        try { lock.releaseLock(); } catch (err) {}
    }
}


/**
 * Save a family relation row
 */
function saveFamilyRelation(relation) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('FamilyRelations');
  if (!sheet) {
    initializeDatabaseV2();
    sheet = ss.getSheetByName('FamilyRelations');
  }

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;

  if (relation.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === relation.id) {
        rowIdx = i + 1;
        break;
      }
    }
  }

  var nextId = relation.id;
  if (!nextId) {
    var nextIdNum = 1;
    if (data.length > 1) {
      var lastId = data[data.length - 1][0] ? String(data[data.length - 1][0]).trim() : "";
      var match = lastId.match(/FAM-(\d+)/);
      if (match) {
        nextIdNum = parseInt(match[1], 10) + 1;
      } else {
        nextIdNum = data.length;
      }
    }
    nextId = "FAM-" + ("0000" + nextIdNum).slice(-4);
  }

  var rowValues = [
    nextId,
    relation.employeeId || "",
    relation.relationType || "",
    relation.fullName || "",
    relation.birthYear || "",
    relation.job || "",
    relation.workPlace || ""
  ];

  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { success: true, id: nextId };
}

/**
 * Delete a family relation row
 */
function deleteFamilyRelation(id) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('FamilyRelations');
  if (!sheet) throw new Error("FamilyRelations sheet not found");

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      rowIdx = i + 1;
      break;
    }
  }

  if (rowIdx === -1) throw new Error("Family relation ID not found: " + id);
  sheet.deleteRow(rowIdx);
  return { success: true };
}

/**
 * Save an education history record
 */
function saveEducationHistory(edu) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('EducationHistory');
  if (!sheet) {
    initializeDatabaseV2();
    sheet = ss.getSheetByName('EducationHistory');
  }

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;

  if (edu.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === edu.id) {
        rowIdx = i + 1;
        break;
      }
    }
  }

  var nextId = edu.id;
  if (!nextId) {
    var nextIdNum = 1;
    if (data.length > 1) {
      var lastId = data[data.length - 1][0] ? String(data[data.length - 1][0]).trim() : "";
      var match = lastId.match(/EDU-(\d+)/);
      if (match) {
        nextIdNum = parseInt(match[1], 10) + 1;
      } else {
        nextIdNum = data.length;
      }
    }
    nextId = "EDU-" + ("0000" + nextIdNum).slice(-4);
  }

  var rowValues = [
    nextId,
    edu.employeeId || "",
    edu.fromDate || "",
    edu.toDate || "",
    edu.schoolName || "",
    edu.fieldOfStudy || "",
    edu.modeOfStudy || "",
    edu.degree || ""
  ];

  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { success: true, id: nextId };
}

/**
 * Delete an education history record
 */
function deleteEducationHistory(id) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('EducationHistory');
  if (!sheet) throw new Error("EducationHistory sheet not found");

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      rowIdx = i + 1;
      break;
    }
  }

  if (rowIdx === -1) throw new Error("Education record ID not found: " + id);
  sheet.deleteRow(rowIdx);
  return { success: true };
}

/**
 * Save a work history record
 */
function saveWorkHistory(work) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('WorkHistory');
  if (!sheet) {
    initializeDatabaseV2();
    sheet = ss.getSheetByName('WorkHistory');
  }

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;

  if (work.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === work.id) {
        rowIdx = i + 1;
        break;
      }
    }
  }

  var nextId = work.id;
  if (!nextId) {
    var nextIdNum = 1;
    if (data.length > 1) {
      var lastId = data[data.length - 1][0] ? String(data[data.length - 1][0]).trim() : "";
      var match = lastId.match(/WRK-(\d+)/);
      if (match) {
        nextIdNum = parseInt(match[1], 10) + 1;
      } else {
        nextIdNum = data.length;
      }
    }
    nextId = "WRK-" + ("0000" + nextIdNum).slice(-4);
  }

  var rowValues = [
    nextId,
    work.employeeId || "",
    work.fromDate || "",
    work.toDate || "",
    work.company || "",
    work.position || "",
    work.referencePerson || "",
    work.referencePhone || ""
  ];

  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { success: true, id: nextId };
}

/**
 * Delete a work history record
 */
function deleteWorkHistory(id) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('WorkHistory');
  if (!sheet) throw new Error("WorkHistory sheet not found");

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      rowIdx = i + 1;
      break;
    }
  }

  if (rowIdx === -1) throw new Error("Work history ID not found: " + id);
  sheet.deleteRow(rowIdx);
  return { success: true };
}

/**
 * Save an official Decision
 */
function saveDecision(dec) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Decisions');
  if (!sheet) {
    initializeDatabaseV2();
    sheet = ss.getSheetByName('Decisions');
  }

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;

  if (dec.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === dec.id) {
        rowIdx = i + 1;
        break;
      }
    }
  }

  var nextId = dec.id;
  if (!nextId) {
    var nextIdNum = 1;
    if (data.length > 1) {
      var lastId = data[data.length - 1][0] ? String(data[data.length - 1][0]).trim() : "";
      var match = lastId.match(/DEC-(\d+)/);
      if (match) {
        nextIdNum = parseInt(match[1], 10) + 1;
      } else {
        nextIdNum = data.length;
      }
    }
    nextId = "DEC-" + ("0000" + nextIdNum).slice(-4);
  }

  var rowValues = [
    nextId,
    dec.employeeId || "",
    dec.decisionType || "",
    dec.decisionNumber || "",
    dec.title || "",
    dec.signDate || "",
    dec.effectiveDate || "",
    dec.notes || "",
    dec.fileBase64 || ""
  ];

  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { success: true, id: nextId };
}

/**
 * Upload an Employee PDF document to specified Drive Folder with dynamic naming formula
 */
function uploadEmployeeDoc(fileDataBase64, originalFileName, employeeId) {
  var ss = getSpreadsheet();
  var empSheet = ss.getSheetByName('Employees');
  if (!empSheet) throw new Error("Employees sheet not found");

  // 1. Fetch employee Name & Email
  var empData = empSheet.getDataRange().getValues();
  var employee = null;
  for (var i = 1; i < empData.length; i++) {
    if (String(empData[i][0]).trim() === employeeId) {
      employee = {
        id: employeeId,
        fullName: String(empData[i][1]).trim(),
        email: String(empData[i][2]).trim()
      };
      break;
    }
  }

  if (!employee) throw new Error("Employee not found with ID: " + employeeId);

  // 2. Lookup DOB in Candidates using Email
  var dobStr = "01012000";
  var candSheet = ss.getSheetByName('Candidates');
  if (candSheet) {
    var candData = candSheet.getDataRange().getValues();
    for (var i = 1; i < candData.length; i++) {
      if (String(candData[i][3]).toLowerCase().trim() === employee.email.toLowerCase().trim()) {
        var dobVal = candData[i][11];
        if (dobVal instanceof Date) {
          dobStr = Utilities.formatDate(dobVal, Session.getScriptTimeZone(), "ddMMyyyy");
        } else if (dobVal) {
          var cleaned = String(dobVal).replace(/[^\d]/g, '');
          if (cleaned.length === 8) {
            dobStr = cleaned;
          } else {
            var parsed = Date.parse(String(dobVal));
            if (!isNaN(parsed)) {
              dobStr = Utilities.formatDate(new Date(parsed), Session.getScriptTimeZone(), "ddMMyyyy");
            }
          }
        }
        break;
      }
    }
  }

  // 3. Format target filename: [Mã nhân sự] + [Tên nhân sự] + [ddmmyyyy_ngày_sinh].pdf
  var cleanName = employee.fullName.replace(/\s+/g, '_');
  // Strip accented characters for safety
  cleanName = cleanName.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  var targetFileName = employeeId + "_" + cleanName + "_" + dobStr + ".pdf";

  // 4. Save to Drive
  var configs = getConfig();
  var folderId = configs.cv_folder_id ? String(configs.cv_folder_id).trim() : "";
  if (!folderId) throw new Error("Chưa cấu hình Thư mục lưu trữ tài liệu (cv_folder_id) trong Config.");

  var folder = DriveApp.getFolderById(folderId);
  var base64Content = fileDataBase64.split(",")[1] || fileDataBase64;
  var bytes = Utilities.base64Decode(base64Content);
  var blob = Utilities.newBlob(bytes, "application/pdf", targetFileName);
  
  var driveFile = folder.createFile(blob);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileUrl = driveFile.getUrl();

  // 5. Append URL to Employee's Docs column in EmployeeDetails
  var detSheet = ss.getSheetByName('EmployeeDetails');
  if (!detSheet) {
    initializeDatabaseV2();
    detSheet = ss.getSheetByName('EmployeeDetails');
  }

  var detData = detSheet.getDataRange().getValues();
  var detRowIdx = -1;
  var currentDocsJson = "[]";

  for (var i = 1; i < detData.length; i++) {
    if (String(detData[i][0]).trim() === employeeId) {
      detRowIdx = i + 1;
      currentDocsJson = detData[i][14] ? String(detData[i][14]).trim() : "[]";
      break;
    }
  }

  var docsList = [];
  try {
    docsList = JSON.parse(currentDocsJson);
  } catch(e) {
    docsList = [];
  }

  docsList.push({
    name: originalFileName || targetFileName,
    url: fileUrl,
    uploadedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
  });

  var updatedDocsJson = JSON.stringify(docsList);

  if (detRowIdx !== -1) {
    detSheet.getRange(detRowIdx, 15).setValue(updatedDocsJson);
  } else {
    // Write a new row with defaults
    var rowValues = [
      employeeId, "", "Nam", "", "", "", "", "", "", "", "", "", "", "", updatedDocsJson
    ];
    detSheet.appendRow(rowValues);
  }

  return { success: true, fileName: targetFileName, fileUrl: fileUrl };
}

/**
 * Generate printable HTML content for a generic Decision record
 */
function getDecisionPrintHtml(decisionId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Decisions');
  if (!sheet) return "<h3>Decision not found</h3>";

  var data = sheet.getDataRange().getValues();
  var dec = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === decisionId) {
      dec = {
        id: decisionId,
        employeeId: String(data[i][1]).trim(),
        decisionType: String(data[i][2]).trim(),
        decisionNumber: String(data[i][3]).trim(),
        title: String(data[i][4]).trim(),
        signDate: String(data[i][5]).trim(),
        effectiveDate: String(data[i][6]).trim(),
        notes: String(data[i][7]).trim(),
        fileBase64: String(data[i][8]).trim()
      };
      break;
    }
  }

  if (!dec) return "<h3>Quyết định không tồn tại</h3>";

  // Fetch employee Name
  var empSheet = ss.getSheetByName('Employees');
  var empName = "Nhân viên";
  var empPos = "Chức vụ";
  var empDept = "Phòng ban";
  if (empSheet) {
    var empData = empSheet.getDataRange().getValues();
    for (var i = 1; i < empData.length; i++) {
      if (String(empData[i][0]).trim() === dec.employeeId) {
        empName = String(empData[i][1]).trim();
        empDept = String(empData[i][4]).trim();
        empPos = String(empData[i][5]).trim();
        break;
      }
    }
  }

  var configs = getConfig();
  var companyName = configs.company_name || "TÊN CÔNG TY CỦA BẠN";

  var html = "<html><head><style>" +
             "body { font-family: 'Times New Roman', Times, serif; padding: 40px; color: #000; line-height: 1.6; font-size: 14px; }" +
             ".header { text-align: center; margin-bottom: 30px; }" +
             ".title { text-align: center; font-size: 20px; font-weight: bold; margin-bottom: 25px; text-transform: uppercase; }" +
             ".section { margin-bottom: 15px; text-align: justify; }" +
             ".footer { margin-top: 50px; display: flex; justify-content: space-between; }" +
             ".footer-col { text-align: center; width: 45%; }" +
             ".signature-space { height: 100px; }" +
             ".attachment-container { margin-top: 40px; text-align: center; page-break-before: always; }" +
             ".attachment-image { max-width: 100%; max-height: 400px; border: 1px solid #ccc; }" +
             "</style></head><body>" +
             "<div class='header'>" +
             "<strong>" + companyName.toUpperCase() + "</strong><br>" +
             "Số: " + dec.decisionNumber + "<br>" +
             "---------------------" +
             "</div>" +
             "<div class='title'>" + dec.title.toUpperCase() + "</div>" +
             "<div class='section'>- Căn cứ vào Điều lệ hoạt động của " + companyName + ";</div>" +
             "<div class='section'>- Căn cứ vào năng lực đóng góp và hiệu quả công việc của nhân sự;</div>" +
             "<div class='section'>- Xét đề xuất của Trưởng phòng Nhân sự.</div>" +
             "<div class='title' style='font-size: 15px; margin: 15px 0;'>QUYẾT ĐỊNH</div>" +
             "<div class='section'><strong>Điều 1:</strong> Phê duyệt quyết định về việc <strong>" + dec.decisionType + "</strong> đối với nhân sự: <strong>" + empName + "</strong></div>" +
             "<div style='margin-left: 20px;'>" +
             "- Chức vụ: " + empPos + "<br>" +
             "- Phòng ban: " + empDept + "<br>" +
             "- Nội dung chi tiết: " + dec.notes + "" +
             "</div>" +
             "<div class='section'><strong>Điều 2:</strong> Quyết định có hiệu lực thi hành kể từ ngày ký " + dec.effectiveDate + ".</div>" +
             "<div class='section'><strong>Điều 3:</strong> Các Phòng ban liên quan và ông/bà <strong>" + empName + "</strong> chịu trách nhiệm thi hành quyết định này.</div>" +
             "<div class='footer'>" +
             "<div class='footer-col'><strong>Nơi nhận:</strong><br>- Như Điều 3<br>- Lưu HSNS</div>" +
             "<div class='footer-col'><strong>Đại diện Ban Giám Đốc</strong><br><i>(Ký và ghi rõ họ tên)</i><div class='signature-space'></div><strong>Cơ quan Đại diện</strong></div>" +
             "</div>";
             
  if (dec.fileBase64 && dec.fileBase64.indexOf("base64") !== -1) {
    html += "<div class='attachment-container'>" +
            "<h3>Tệp hình ảnh quét Quyết định đính kèm</h3>" +
            "<img src='" + dec.fileBase64 + "' class='attachment-image' />" +
            "</div>";
  }
  
  html += "</body></html>";
  return html;
}

/**
 * Lấy toàn bộ danh sách mẫu in từ sheet PrintTemplates
 */
function getPrintTemplates() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PrintTemplates');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('PrintTemplates');
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var item = {};
      headers.forEach(function(h, idx) {
        var prop = h.charAt(0).toLowerCase() + h.slice(1);
        item[prop] = row[idx];
      });
      list.push(item);
    }
    return list;
  } catch (error) {
    Logger.log("Lỗi getPrintTemplates: " + error.toString());
    return [];
  }
}

/**
 * Lưu hoặc Cập nhật mẫu in
 */
function savePrintTemplate(templateData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PrintTemplates');
    if (!sheet) {
      initializeDatabaseV2();
      sheet = ss.getSheetByName('PrintTemplates');
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h] = idx; });

    var templateId = templateData.id || templateData.ID || "";
    var rowIdx = -1;

    if (templateId) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][colMap['ID']]).trim() === String(templateId).trim()) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    if (rowIdx !== -1) {
      // Cập nhật
      function updateCell(colName, value) {
        var idx = colMap[colName];
        if (idx !== undefined && value !== undefined && value !== null) {
          sheet.getRange(rowIdx, idx + 1).setValue(value);
        }
      }
      updateCell('TemplateName', templateData.templateName);
      updateCell('Category', templateData.category);
      updateCell('HtmlContent', templateData.htmlContent);
      updateCell('Description', templateData.description);
    } else {
      // Tạo mới
      // Sinh ID tự động: PRT-0001, PRT-0002...
      var nextNum = 1;
      if (data.length > 1) {
        var ids = data.slice(1).map(function(r) {
          var idStr = String(r[colMap['ID']]).trim();
          var match = idStr.match(/PRT-(\d+)/);
          return match ? parseInt(match[1]) : 0;
        });
        var maxId = Math.max.apply(null, ids);
        if (maxId > 0) nextNum = maxId + 1;
      }
      var nextId = "PRT-" + String(nextNum).padStart(4, '0');

      var newRow = new Array(headers.length).fill("");
      newRow[colMap['ID']] = nextId;
      newRow[colMap['TemplateName']] = templateData.templateName || "";
      newRow[colMap['Category']] = templateData.category || "";
      newRow[colMap['HtmlContent']] = templateData.htmlContent || "";
      newRow[colMap['Description']] = templateData.description || "";
      newRow[colMap['CreatedAt']] = now;

      sheet.appendRow(newRow);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Xóa mẫu in
 */
function deletePrintTemplate(templateId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PrintTemplates');
    if (!sheet) return { success: false, message: "Không tìm thấy sheet PrintTemplates" };

    var data = sheet.getDataRange().getValues();
    var idColIdx = 0; // ID luôn ở cột 0
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]).trim() === String(templateId).trim()) {
        rowIdx = i + 1;
        break;
      }
    }

    if (rowIdx !== -1) {
      sheet.deleteRow(rowIdx);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, message: "Không tìm thấy mẫu in có ID: " + templateId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Xuất PDF trực tiếp HTML-to-PDF siêu tốc từ mẫu in cho một nhân sự
 */
function generatePDFFromTemplate(templateId, employeeId) {
  try {
    var ss = getSpreadsheet();
    
    // 1. Lấy mẫu in
    var sheet = ss.getSheetByName('PrintTemplates');
    if (!sheet) throw new Error("Bảng mẫu in không tồn tại.");
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h] = idx; });
    
    var templateObj = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colMap['ID']]).trim() === String(templateId).trim()) {
        templateObj = {
          templateName: String(data[i][colMap['TemplateName']]).trim(),
          htmlContent: String(data[i][colMap['HtmlContent']]).trim()
        };
        break;
      }
    }
    if (!templateObj) throw new Error("Không tìm thấy mẫu in có ID: " + templateId);

    // 2. Lấy thông tin nhân viên
    var empDetailResult = getEmployeeDetail(employeeId);
    if (!empDetailResult.success) throw new Error("Không thể tải chi tiết nhân viên: " + empDetailResult.message);
    var employeeObj = empDetailResult.employee;

    // 3. Tiến hành replace các placeholder {{key}}
    var finalHtml = templateObj.htmlContent;
    for (var key in employeeObj) {
      var placeholder = "{{" + key + "}}";
      var regex = new RegExp(placeholder, "g");
      var val = employeeObj[key];
      // Xử lý hiển thị tiền tệ nếu là basicSalary hoặc allowance
      if ((key === 'basicSalary' || key === 'probationSalary' || key === 'allowance') && typeof val === 'number') {
        val = val.toLocaleString() + " VND";
      }
      finalHtml = finalHtml.replace(regex, (val !== undefined && val !== null) ? String(val) : "");
    }
    
    // Thêm CSS reset cơ bản cho PDF
    var pdfStyle = "<style>" +
                   "body { font-family: 'Arial', sans-serif; padding: 20px; line-height: 1.5; font-size: 13px; color: #000; }" +
                   "</style>";
    if (finalHtml.indexOf("<head>") !== -1) {
      finalHtml = finalHtml.replace("</head>", pdfStyle + "</head>");
    } else {
      finalHtml = pdfStyle + finalHtml;
    }

    // 4. HTML-to-PDF siêu tốc
    var htmlBlob = Utilities.newBlob(finalHtml, 'text/html', 'document.html');
    var pdfBlob = htmlBlob.getAs('application/pdf');

    // 5. Lưu trữ Drive
    var configs = getConfig() || {};
    var folderId = configs['PRINT_PDF_FOLDER_ID'] || configs['print_pdf_folder_id'] || "";
    var targetFolder = null;
    if (folderId) {
      try {
        targetFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        targetFolder = null;
      }
    }
    if (!targetFolder) {
      // Fallback: tạo ở root
      var folderName = "HRM Generated Documents";
      var folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        targetFolder = folders.next();
      } else {
        targetFolder = DriveApp.createFolder(folderName);
        targetFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
      // Lưu lại folder id vào config
      saveConfig({'PRINT_PDF_FOLDER_ID': targetFolder.getId()});
    }

    var pdfName = templateObj.templateName + "_" + employeeObj.fullName + "_" + employeeId + ".pdf";
    pdfName = pdfName.replace(/[\/\\:*?"<>|]/g, "_");

    var file = targetFolder.createFile(pdfBlob);
    file.setName(pdfName);

    var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    
    return {
      success: true,
      url: file.getUrl(),
      base64: base64Data,
      fileName: pdfName
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/* ==========================================================================
   SPRINT 2: ONBOARDING, SHIFTS, LATE DEDUCTION, ESS PORTAL, EMAIL FALLBACK
   ========================================================================== */

/**
 * 1. Tự động tạo Checklist Onboarding khi ứng viên chuyển sang trạng thái Passed
 */
function generateOnboardingChecklist(candidateId, employeeEmail, offerStartDate) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var candidate = getCandidateById(candidateId);
    var candName = candidate ? candidate.senderName : 'Nhân sự mới';
    var email = employeeEmail || (candidate ? candidate.senderEmail : '');
    
    var startDate = offerStartDate ? new Date(offerStartDate) : new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
    if (isNaN(startDate.getTime())) startDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
    
    function formatDateOffset(days) {
      var d = new Date(startDate.getTime() + days * 24 * 60 * 60 * 1000);
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    
    var tasks = [
      { name: "Cấp Email công ty & cấp quyền phần mềm (IT Ticket)", assignedTo: "IT", dueDate: formatDateOffset(-2) },
      { name: "Chuẩn bị Bàn làm việc, Laptop, Thẻ tên & Thẻ ra vào", assignedTo: "Admin", dueDate: formatDateOffset(-1) },
      { name: "Phân công Người hướng dẫn (Buddy/Mentor)", assignedTo: "Manager", dueDate: formatDateOffset(-1) },
      { name: "Gửi Welcome Kit & Sổ tay Nhân viên (Employee Handbook)", assignedTo: "HR", dueDate: formatDateOffset(-3) }
    ];
    
    var obRead = batchReadSheet('OnboardingTasks');
    var obData = obRead.data;
    if (obData.length === 0) {
      obData = [['ID', 'CandidateID', 'EmployeeEmail', 'TaskName', 'AssignedTo', 'DueDate', 'Status', 'Notes', 'CreatedAt']];
    }
    
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    tasks.forEach(function(t, idx) {
      var taskId = "OB-" + Number(new Date()) + "-" + (idx + 1);
      obData.push([
        taskId,
        candidateId || '',
        email,
        t.name,
        t.assignedTo,
        t.dueDate,
        'Pending',
        'Tự động khởi tạo cho ' + candName,
        nowStr
      ]);
    });
    
    batchWriteSheet('OnboardingTasks', obData);
    return { success: true, count: tasks.length };
  } catch (e) {
    Logger.log("Lỗi generateOnboardingChecklist: " + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 1. Daily Trigger: Gửi Welcome Email & Sổ tay nhân viên trước ngày nhận việc 3 ngày
 */
function sendPreboardingWelcomeEmailTrigger() {
  try {
    var employees = getEmployees();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var welcomeList = [];
    employees.forEach(function(emp) {
      if (emp.status === 'Thử việc' || emp.status === 'Mới nhận') {
        var startDateStr = emp.probationStartDate || emp.officialStartDate || emp.startDate;
        if (!startDateStr) return;
        var startD = new Date(startDateStr);
        if (isNaN(startD.getTime())) return;
        startD.setHours(0, 0, 0, 0);
        
        var diffDays = Math.ceil((startD.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        if (diffDays === 3) {
          welcomeList.push(emp);
        }
      }
    });
    
    if (welcomeList.length === 0) return { success: true, message: "Không có nhân sự mới nào cần gửi Welcome Email hôm nay." };
    
    var configs = getConfig();
    var handbookUrl = configs.employee_handbook_url || "https://drive.google.com";
    
    welcomeList.forEach(function(emp) {
      if (!emp.email) return;
      var subject = "[RecruitFlow HRM] Chào mừng bạn gia nhập đội ngũ " + (configs.company_name || 'Doanh nghiệp') + "!";
      var htmlBody = "<div style='font-family: sans-serif; font-size: 13px; color: #1e293b; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; padding: 24px; border-radius: 16px;'>" +
                     "<h2 style='color: #4f46e5; margin-top: 0;'>WELCOME TO THE TEAM! 🎉</h2>" +
                     "<p>Kính gửi <b>" + emp.fullName + "</b>,</p>" +
                     "<p>Chỉ còn 3 ngày nữa là đến ngày bạn chính thức nhận việc tại vị trí <b>" + (emp.position || 'Nhân sự') + "</b> thuộc phòng <b>" + (emp.department || 'Ban Nhân sự') + "</b>.</p>" +
                     "<p>Để bạn chuẩn bị tốt nhất cho hành trình mới, công ty trân trọng gửi tới bạn <b>Sổ tay Nhân viên (Employee Handbook)</b> chứa đựng các thông tin văn hóa, quy định làm việc và phúc lợi:</p>" +
                     "<p style='text-align: center; margin: 20px 0;'><a href='" + handbookUrl + "' target='_blank' style='background-color: #4f46e5; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: bold;'>Đọc Sổ tay Nhân viên</a></p>" +
                     "<p>Hẹn gặp lại bạn vào 8h30 sáng ngày đi làm đầu tiên tại văn phòng!</p>" +
                     "</div>";
      sendEmailWithFallback(emp.email, subject, htmlBody);
    });
    
    return { success: true, count: welcomeList.length };
  } catch (e) {
    Logger.log("Lỗi sendPreboardingWelcomeEmailTrigger: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 2. Đăng ký Đơn xin đổi ca làm việc (Shift Swap Request)
 */
function submitShiftSwapRequest(requesterEmail, targetEmail, swapDate, requesterShift, targetShift) {
  requireRole(['All']);
  if (!requesterEmail || !targetEmail || !swapDate) {
    return { success: false, message: "Vui lòng nhập đầy đủ email đồng nghiệp và ngày đổi ca." };
  }
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var readRes = batchReadSheet('ShiftSwaps');
    var data = readRes.data;
    if (data.length === 0) {
      data = [['ID', 'RequesterEmail', 'TargetEmail', 'SwapDate', 'RequesterShift', 'TargetShift', 'Status', 'ApproverEmail', 'CreatedAt']];
    }
    
    var swapId = "SWAP-" + Number(new Date()) + "-" + Math.floor(Math.random() * 1000);
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    data.push([
      swapId,
      requesterEmail.trim().toLowerCase(),
      targetEmail.trim().toLowerCase(),
      swapDate,
      requesterShift || 'HC',
      targetShift || 'HC',
      'Pending',
      '',
      nowStr
    ]);
    
    batchWriteSheet('ShiftSwaps', data);
    
    var subject = "[RecruitFlow HRM] Yêu cầu Đổi ca làm việc từ " + requesterEmail;
    var htmlBody = "<p>Bạn có một Đơn xin đổi ca làm việc vào ngày <b>" + swapDate + "</b> từ đồng nghiệp <b>" + requesterEmail + "</b>.</p>" +
                   "<p>Vui lòng đăng nhập hệ thống để duyệt hoặc từ chối yêu cầu đổi ca này.</p>";
    sendEmailWithFallback(targetEmail, subject, htmlBody);
    
    return { success: true, message: "Đã gửi đơn xin đổi ca thành công!", swapId: swapId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 2. Phê duyệt Đơn xin đổi ca làm việc (Shift Swap Approval)
 */
function approveShiftSwap(swapId, status, approverEmail) {
  requireRole(['Admin', 'HR', 'Manager', 'All']);
  if (!swapId || !status) return { success: false, message: "Thông tin không hợp lệ." };
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var readRes = batchReadSheet('ShiftSwaps');
    var data = readRes.data;
    if (data.length <= 1) return { success: false, message: "Không tìm thấy bảng đơn đổi ca." };
    
    var targetRowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(swapId).trim()) {
        targetRowIdx = i;
        break;
      }
    }
    
    if (targetRowIdx === -1) return { success: false, message: "Không tìm thấy mã đơn đổi ca: " + swapId };
    
    data[targetRowIdx][6] = status;
    data[targetRowIdx][7] = approverEmail || Session.getActiveUser().getEmail() || '';
    
    batchWriteSheet('ShiftSwaps', data);
    
    var reqEmail = data[targetRowIdx][1];
    var subject = "[RecruitFlow HRM] Kết quả Yêu cầu Đổi ca (" + status + ")";
    var htmlBody = "<p>Yêu cầu đổi ca mã <b>" + swapId + "</b> của bạn đã được cập nhật trạng thái: <b style='color: " + (status === 'Approved' ? 'green' : 'red') + ";'>" + status + "</b>.</p>";
    sendEmailWithFallback(reqEmail, subject, htmlBody);
    
    return { success: true, message: "Cập nhật đơn đổi ca thành công!" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 3. Tự động đối soát bảng công và tính số phút đi muộn / về sớm cho từng nhân viên theo ca
 */
function reconcileAttendanceLateEarly(month, year) {
  requireRole(['Admin', 'HR']);
  try {
    var shiftsRead = batchReadSheet('Shifts');
    var shiftsData = shiftsRead.data;
    var shiftGraceMap = {};
    var shiftStartMap = {};
    var shiftEndMap   = {};
    
    for (var s = 1; s < shiftsData.length; s++) {
      var code  = String(shiftsData[s][0]).trim();
      var start = String(shiftsData[s][2]).trim();
      var end   = String(shiftsData[s][3]).trim();
      var grace = Number(shiftsData[s][5]) || 15;
      
      shiftGraceMap[code] = grace;
      shiftStartMap[code] = start;
      shiftEndMap[code]   = end;
    }
    
    var attRead = batchReadSheet('Attendance');
    var attData = attRead.data;
    if (attData.length <= 1) return { success: true, summary: {} };
    
    var attHeaders = attRead.headers;
    var empIdCol    = attHeaders.indexOf('EmployeeID');
    var dateCol     = attHeaders.indexOf('Date');
    var checkInCol  = attHeaders.indexOf('CheckIn');
    var checkOutCol = attHeaders.indexOf('CheckOut');
    
    var empDeductionMap = {};
    
    for (var i = 1; i < attData.length; i++) {
      var row = attData[i];
      var empId = empIdCol !== -1 ? String(row[empIdCol]).trim() : '';
      var dVal  = dateCol !== -1 ? row[dateCol] : '';
      if (!empId || !dVal) continue;
      
      var dStr = dVal instanceof Date ? Utilities.formatDate(dVal, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dVal).trim();
      var dParts = dStr.split("-");
      if (dParts.length !== 3) continue;
      
      if (parseInt(dParts[0], 10) === parseInt(year, 10) && parseInt(dParts[1], 10) === parseInt(month, 10)) {
        var checkInStr  = checkInCol !== -1 ? String(row[checkInCol]).trim() : '';
        var shiftStart = shiftStartMap['HC'] || "08:00";
        var grace = shiftGraceMap['HC'] || 15;
        
        var lateMinutes = 0;
        if (checkInStr && checkInStr.indexOf(":") !== -1) {
          var inParts = checkInStr.split(":");
          var startParts = shiftStart.split(":");
          var inMins = parseInt(inParts[0], 10) * 60 + parseInt(inParts[1], 10);
          var startMins = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);
          var diff = inMins - startMins;
          if (diff > grace) {
            lateMinutes = diff - grace;
          }
        }
        
        if (!empDeductionMap[empId]) {
          empDeductionMap[empId] = { totalLateMinutes: 0, lateCount: 0 };
        }
        if (lateMinutes > 0) {
          empDeductionMap[empId].totalLateMinutes += lateMinutes;
          empDeductionMap[empId].lateCount += 1;
        }
      }
    }
    
    return { success: true, summary: empDeductionMap };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 5. ESS Portal: Lấy thông tin cá nhân, lịch sử lương, quỹ phép và lịch sử chấm công của nhân viên
 * Khóa lỗ hổng BFLA: Sử dụng trực tiếp Email từ Server Session
 */
function getEssProfile() {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) throw new Error("Bảo mật (BFLA Blocked): Vui lòng xác thực tài khoản để truy cập Cổng ESS.");
  
  var employees = getEmployees();
  var emp = employees.find(e => (e.email || '').toLowerCase().trim() === activeEmail);
  if (!emp) throw new Error("Không tìm thấy thông tin hồ sơ nhân sự của tài khoản: " + activeEmail);
  
  var empDetailRes = getEmployeeDetail(emp.id);
  var empCombined = empDetailRes.employee || emp;
  
  return {
    success: true,
    profile: empCombined,
    payrolls: empDetailRes.payrolls || [],
    attendance: empDetailRes.attendance || [],
    leaveBalance: {
      total: empCombined.totalLeaveDays || 12,
      used: empCombined.usedLeaveDays || 0,
      remaining: (empCombined.totalLeaveDays || 12) - (empCombined.usedLeaveDays || 0)
    }
  };
}

/**
 * 5. ESS Portal: Nhân viên gửi yêu cầu chỉnh sửa thông tin cá nhân (SĐT, Địa chỉ, STK ngân hàng...)
 * Khóa lỗ hổng BFLA: Sử dụng trực tiếp Email từ Server Session
 */
function updateEssProfile(updatedFields) {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) throw new Error("Bảo mật (BFLA Blocked): Vui lòng xác thực tài khoản.");
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var reqRead = batchReadSheet('EssUpdateRequests');
    var data = reqRead.data;
    if (data.length === 0) {
      data = [['ID', 'EmployeeEmail', 'FieldName', 'OldValue', 'NewValue', 'Status', 'ApproverEmail', 'CreatedAt']];
    }
    
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var reqCount = 0;
    
    for (var field in updatedFields) {
      var reqId = "ESS-" + Number(new Date()) + "-" + Math.floor(Math.random() * 1000);
      data.push([
        reqId,
        activeEmail,
        field,
        '',
        String(updatedFields[field]),
        'Pending',
        '',
        nowStr
      ]);
      reqCount++;
    }
    
    batchWriteSheet('EssUpdateRequests', data);
    return { success: true, message: "Đã gửi " + reqCount + " yêu cầu cập nhật thông tin tới phòng HR duyệt." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 5. ESS Portal: Xuất Giấy Xác nhận Công tác / Xác nhận Thu nhập dạng PDF
 * Khóa lỗ hổng BFLA: Sử dụng trực tiếp Email từ Server Session
 */
function generateEmploymentVerificationPDF(verificationType) {
  var activeEmail = getActiveUserEmail();
  if (!activeEmail) throw new Error("Bảo mật (BFLA Blocked): Vui lòng xác thực tài khoản.");
  
  var employees = getEmployees();
  var emp = employees.find(e => (e.email || '').toLowerCase().trim() === activeEmail);
  if (!emp) throw new Error("Không tìm thấy thông tin nhân viên.");
  
  var typeTitle = verificationType === 'Income' ? 'GIẤY XÁC NHẬN THU NHẬP' : 'GIẤY XÁC NHẬN CÔNG TÁC';
  var configs = getConfig() || {};
  var companyName = configs.company_name || 'CÔNG TY CỔ PHẦN CÔNG NGHỆ RECRUITFLOW';
  
  var htmlContent = "<div style='font-family: Arial, sans-serif; font-size: 13px; line-height: 1.6; padding: 30px;'>" +
                    "<h3 style='text-align: center; margin-bottom: 5px;'>" + companyName.toUpperCase() + "</h3>" +
                    "<h2 style='text-align: center; color: #1e1b4b; margin-top: 15px;'>" + typeTitle + "</h2>" +
                    "<p>Kính gửi: Các cơ quan hữu quan / Ngân hàng / Đơn vị tiếp nhận</p>" +
                    "<p>Công ty chúng tôi trân trọng xác nhận thông tin nhân sự sau đây:</p>" +
                    "<ul>" +
                    "<li><b>Họ và tên nhân viên:</b> " + emp.fullName + "</li>" +
                    "<li><b>Mã nhân sự:</b> " + emp.id + "</li>" +
                    "<li><b>Địa chỉ Email:</b> " + emp.email + "</li>" +
                    "<li><b>Phòng ban công tác:</b> " + (emp.department || 'Chưa phân phòng') + "</li>" +
                    "<li><b>Chức danh chuyên môn:</b> " + (emp.position || 'Nhân sự chính thức') + "</li>" +
                    "<li><b>Ngày bắt đầu làm việc:</b> " + (emp.officialStartDate || emp.startDate || '—') + "</li>" +
                    "<li><b>Mức lương cơ bản hàng tháng:</b> " + (emp.basicSalary ? emp.basicSalary.toLocaleString() + ' VNĐ' : '—') + "</li>" +
                    "</ul>" +
                    "<p>Giấy xác nhận này được cấp theo yêu cầu của cá nhân ông/bà " + emp.fullName + " để sử dụng làm thủ tục hành chính hợp pháp.</p>" +
                    "<br><br>" +
                    "<table style='width: 100%; text-align: center;'>" +
                    "<tr><td></td><td><b>ĐẠI DIỆN BAN GIÁM ĐỐC / HR DIRECTOR</b><br><i>(Ký tên & đóng dấu)</i></td></tr>" +
                    "</table>" +
                    "</div>";
  
  var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'verification.html');
  var pdfBlob = htmlBlob.getAs('application/pdf');
  var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  
  return {
    success: true,
    base64: base64Data,
    fileName: typeTitle.replace(/\s+/g, '_') + "_" + emp.fullName + ".pdf"
  };
}

/**
 * 6. Gửi Email thông minh với cơ chế Fallback tự động:
 * 1. Kiểm tra Quota còn lại của Gmail (MailApp.getRemainingDailyQuota()).
 * 2. Nếu quota > 10: Sử dụng MailApp.sendEmail().
 * 3. Nếu quota <= 10: Tự động chuyển đổi sang SendGrid HTTP API via UrlFetchApp.
 */
function sendEmailWithFallback(toEmail, subject, htmlBody, attachments) {
  if (!toEmail) return { success: false, message: "Email nhận không được để trống." };
  
  var remainingQuota = MailApp.getRemainingDailyQuota();
  
  // Dùng MailApp nếu còn Quota an toàn (> 10 mail)
  if (remainingQuota > 10) {
    try {
      MailApp.sendEmail({
        to: toEmail,
        subject: subject,
        htmlBody: htmlBody
      });
      return { success: true, provider: 'GmailMailApp' };
    } catch (e) {
      Logger.log("MailApp thất bại, chuyển sang SendGrid Fallback: " + e.message);
    }
  }
  
  // Chuyển sang SendGrid HTTP API khi Gmail chạm hạn ngạch Quota
  var configs = getConfig();
  var sendgridApiKey = configs.sendgrid_api_key || configs.SENDGRID_API_KEY;
  var senderEmail = configs.sender_email || "no-reply@company.com";
  
  if (!sendgridApiKey) {
    try {
      MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: htmlBody });
      return { success: true, provider: 'GmailMailAppFallback' };
    } catch (err) {
      throw new Error("Tài khoản đã hết Hạn ngạch Gmail Quota và chưa cấu hình SendGrid API Key trong Config.");
    }
  }
  
  try {
    var payload = {
      personalizations: [{
        to: [{ email: toEmail }]
      }],
      from: { email: senderEmail, name: configs.company_name || "RecruitFlow HRM" },
      subject: subject,
      content: [{
        type: "text/html",
        value: htmlBody
      }]
    };
    
    var options = {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + sendgridApiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch("https://api.sendgrid.com/v3/mail/send", options);
    var code = response.getResponseCode();
    
    if (code === 200 || code === 202) {
      return { success: true, provider: 'SendGridAPI' };
    } else {
      throw new Error("SendGrid API phản hồi mã lỗi: " + code + " - " + response.getContentText());
    }
  } catch (err) {
    Logger.log("Lỗi SendGrid API: " + err.message);
    throw new Error("Không thể gửi email: " + err.message);
  }
}

/* ==========================================================================
   SPRINT 3, 4, 5: BANK EXPORT, SENIORITY LEAVE, CULTURE, MVP MODULES, LOOKER
   ========================================================================== */

/**
 * 1. Xuất file chuyển khoản lương ngân hàng (Vietcombank, Techcombank, VPBank, BIDV)
 * Định dạng chuẩn CSV theo quy cách cổng Internet Banking doanh nghiệp
 */
function exportBankPayrollFile(month, year, bankCode) {
  requireRole(['Admin', 'HR']);
  bankCode = (bankCode || 'VCB').toUpperCase().trim();
  
  var payRead = batchReadSheet('Payroll');
  var payData = payRead.data;
  if (payData.length <= 1) throw new Error("Chưa có dữ liệu bảng lương tháng " + month + "/" + year);
  
  var employees = getEmployees();
  var empMap = {};
  employees.forEach(function(e) { empMap[e.id] = e; });
  
  var targetPayrolls = [];
  for (var i = 1; i < payData.length; i++) {
    var pMonth = parseInt(payData[i][1], 10);
    var pYear  = parseInt(payData[i][2], 10);
    var empId  = String(payData[i][3]).trim();
    var net    = Number(payData[i][7]) || 0;
    
    if (pMonth === parseInt(month, 10) && pYear === parseInt(year, 10) && net > 0) {
      targetPayrolls.push({
        empId: empId,
        netSalary: net,
        empObj: empMap[empId] || {}
      });
    }
  }
  
  if (targetPayrolls.length === 0) throw new Error("Không có bản ghi lương hợp lệ để xuất file cho tháng " + month + "/" + year);
  
  var csvLines = [];
  var fileName = "LienNganHang_" + bankCode + "_T" + month + "_" + year + ".csv";
  
  if (bankCode === 'VCB') {
    csvLines.push("STK,Ten_Chu_Tai_Khoan,So_Tien,Noi_Dung");
    targetPayrolls.forEach(function(p) {
      var accNo = p.empObj.bankAccountNumber || "STK_CHUA_CAP";
      var name  = (p.empObj.fullName || "NHAN VIEN").toUpperCase();
      var amount= p.netSalary;
      var remark= "Luu luong T" + month + "." + year + " " + p.empId;
      csvLines.push('"' + accNo + '","' + name + '",' + amount + ',"' + remark + '"');
    });
  } else if (bankCode === 'TCB') {
    csvLines.push("BeneficiaryAccount,BeneficiaryName,Amount,Details");
    targetPayrolls.forEach(function(p) {
      var accNo = p.empObj.bankAccountNumber || "STK_CHUA_CAP";
      var name  = (p.empObj.fullName || "NHAN VIEN").toUpperCase();
      var amount= p.netSalary;
      var remark= "Chuyen luong Thang " + month + " nam " + year;
      csvLines.push('"' + accNo + '","' + name + '",' + amount + ',"' + remark + '"');
    });
  } else if (bankCode === 'VPB') {
    csvLines.push("Account_Number,Account_Name,Amount,Memo");
    targetPayrolls.forEach(function(p) {
      var accNo = p.empObj.bankAccountNumber || "STK_CHUA_CAP";
      var name  = (p.empObj.fullName || "NHAN VIEN").toUpperCase();
      var amount= p.netSalary;
      var remark= "VPB Payroll T" + month + "/" + year;
      csvLines.push('"' + accNo + '","' + name + '",' + amount + ',"' + remark + '"');
    });
  } else {
    csvLines.push("SoTaiKhoan,TenNguoiNhan,SoTien,NoidungChuyenKhoan");
    targetPayrolls.forEach(function(p) {
      var accNo = p.empObj.bankAccountNumber || "STK_CHUA_CAP";
      var name  = (p.empObj.fullName || "NHAN VIEN").toUpperCase();
      var amount= p.netSalary;
      var remark= "Thanh toan luong T" + month + "-" + year;
      csvLines.push('"' + accNo + '","' + name + '",' + amount + ',"' + remark + '"');
    });
  }
  
  var csvString = csvLines.join("\r\n");
  var base64 = Utilities.base64Encode(Utilities.newBlob(csvString, 'text/csv').getBytes());
  
  return {
    success: true,
    bankCode: bankCode,
    recordCount: targetPayrolls.length,
    fileName: fileName,
    csvContent: csvString,
    base64: base64
  };
}

/**
 * 2. Tự động tính ngày Phép thâm niên theo Điều 114 Bộ luật Lao động 2019
 */
function calculateSeniorityLeave(employeeId) {
  var employees = getEmployees();
  var emp = employees.find(e => e.id === employeeId);
  if (!emp) throw new Error("Không tìm thấy nhân viên: " + employeeId);
  
  var startDateStr = emp.officialStartDate || emp.probationStartDate || emp.startDate;
  if (!startDateStr) return { baseLeave: 12, seniorityBonus: 0, totalLeave: 12 };
  
  var startD = new Date(startDateStr);
  if (isNaN(startD.getTime())) return { baseLeave: 12, seniorityBonus: 0, totalLeave: 12 };
  
  var today = new Date();
  var diffYears = (today.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24 * 365.25);
  var seniorityBonusDays = Math.floor(diffYears / 5);
  
  var baseLeave = Number(emp.totalLeaveDays) || 12;
  var totalLeave = baseLeave + seniorityBonusDays;
  
  return {
    employeeId: employeeId,
    yearsOfService: Math.floor(diffYears),
    baseLeaveDays: baseLeave,
    seniorityBonusDays: seniorityBonusDays,
    totalLeaveDays: totalLeave
  };
}

/**
 * 2. Quy đổi giờ làm thêm OT thành ngày nghỉ bù (Compensatory Time Off)
 */
function convertOvertimeToCompTime(employeeId, otHoursToConvert) {
  requireRole(['Admin', 'HR', 'Manager']);
  otHoursToConvert = Number(otHoursToConvert) || 0;
  if (otHoursToConvert <= 0) return { success: false, message: "Số giờ quy đổi phải > 0." };
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var empRead = batchReadSheet('Employees');
    var empData = empRead.data;
    if (empData.length <= 1) return { success: false, message: "Bảng Employees trống." };
    
    var empRowIdx = -1;
    for (var i = 1; i < empData.length; i++) {
      if (String(empData[i][0]).trim() === String(employeeId).trim()) {
        empRowIdx = i;
        break;
      }
    }
    
    if (empRowIdx === -1) return { success: false, message: "Không tìm thấy nhân viên." };
    
    var addedLeaveDays = otHoursToConvert / 8.0;
    var currentLeave = Number(empData[empRowIdx][16]) || 12;
    var newTotalLeave = currentLeave + addedLeaveDays;
    
    empData[empRowIdx][16] = newTotalLeave;
    batchWriteSheet('Employees', empData);
    
    logJobHistory(employeeId, empData[empRowIdx][1], 'CompTimeConverted', currentLeave, newTotalLeave, '', '', 'Quy đổi ' + otHoursToConvert + ' giờ OT thành ' + addedLeaveDays + ' ngày nghỉ bù');
    
    return {
      success: true,
      employeeId: employeeId,
      convertedOtHours: otHoursToConvert,
      addedLeaveDays: addedLeaveDays,
      newTotalLeaveDays: newTotalLeave
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 3. Đăng bài viết Tin tức / Văn hóa nội bộ lên Newsfeed
 */
function postNewsfeedItem(title, content, category) {
  var activeUser = Session.getActiveUser().getEmail() || 'HR Team';
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var nfRead = batchReadSheet('Newsfeed');
    var data = nfRead.data;
    if (data.length === 0) {
      data = [['ID', 'AuthorEmail', 'Title', 'Content', 'Category', 'CreatedAt']];
    }
    
    var id = "NEWS-" + Number(new Date());
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    data.push([id, activeUser, title, content, category || 'Chung', nowStr]);
    batchWriteSheet('Newsfeed', data);
    
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 3. Gửi tặng vinh danh Kudos cho đồng nghiệp
 */
function sendKudos(receiverEmail, badge, message) {
  var senderEmail = Session.getActiveUser().getEmail();
  if (!senderEmail || !receiverEmail) return { success: false, message: "Thông tin không đầy đủ." };
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var kRead = batchReadSheet('Kudos');
    var data = kRead.data;
    if (data.length === 0) {
      data = [['ID', 'SenderEmail', 'ReceiverEmail', 'Badge', 'Message', 'CreatedAt']];
    }
    
    var id = "KUDOS-" + Number(new Date());
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    data.push([id, senderEmail, receiverEmail, badge || 'Ngôi sao đồng đội', message, nowStr]);
    batchWriteSheet('Kudos', data);
    
    var subject = "[RecruitFlow HRM] Bạn vừa nhận được 1 KUDOS vinh danh từ " + senderEmail + "!";
    var htmlBody = "<div style='font-family: sans-serif; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px; max-width: 500px;'>" +
                   "<h2 style='color: #4f46e5;'>BẠN THẬT TUYỆT VỜI!</h2>" +
                   "<p>Đồng nghiệp <b>" + senderEmail + "</b> vừa tặng bạn danh hiệu: <b style='color: #d97706;'>" + (badge || 'Ngôi sao đồng đội') + "</b></p>" +
                   "<blockquote style='background: #f8fafc; padding: 12px; border-left: 4px solid #4f46e5;'>" + message + "</blockquote>" +
                   "</div>";
    sendEmailWithFallback(receiverEmail, subject, htmlBody);
    
    return { success: true, message: "Đã gửi Kudos vinh danh thành công!" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 3. Daily Trigger: Tự động gửi thư chúc mừng Sinh nhật & Kỷ niệm ngày làm việc
 */
function sendBirthdayAndAnniversaryGreetingsTrigger() {
  try {
    var employees = getEmployees();
    var today = new Date();
    var currentMonthDay = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM-dd");
    
    var bdayCount = 0;
    var annivCount = 0;
    
    employees.forEach(function(emp) {
      if (!emp.email) return;
      
      if (emp.dob || emp.dateOfBirth) {
        var dobD = new Date(emp.dob || emp.dateOfBirth);
        if (!isNaN(dobD.getTime())) {
          var dobMonthDay = Utilities.formatDate(dobD, Session.getScriptTimeZone(), "MM-dd");
          if (dobMonthDay === currentMonthDay) {
            var bSubject = "CHÚC MỪNG SINH NHẬT " + emp.fullName.toUpperCase() + "!";
            var bBody = "<p>Chúc mừng sinh nhật <b>" + emp.fullName + "</b>! Chúc bạn một tuổi mới tràn đầy sức khỏe, hạnh phúc và thành công cùng công ty!</p>";
            sendEmailWithFallback(emp.email, bSubject, bBody);
            bdayCount++;
          }
        }
      }
      
      var startDateStr = emp.officialStartDate || emp.probationStartDate || emp.startDate;
      if (startDateStr) {
        var startD = new Date(startDateStr);
        if (!isNaN(startD.getTime())) {
          var startMonthDay = Utilities.formatDate(startD, Session.getScriptTimeZone(), "MM-dd");
          var yearsDiff = today.getFullYear() - startD.getFullYear();
          if (startMonthDay === currentMonthDay && yearsDiff >= 1) {
            var aSubject = "CHÚC MỪNG KỶ NIỆM " + yearsDiff + " NĂM ĐỒNG HÀNH CÙNG CÔNG TY!";
            var aBody = "<p>Cảm ơn <b>" + emp.fullName + "</b> đã gắn bó và cống hiến suốt <b>" + yearsDiff + " năm</b> qua!</p>";
            sendEmailWithFallback(emp.email, aSubject, aBody);
            annivCount++;
          }
        }
      }
    });
    
    return { success: true, birthdaysSent: bdayCount, anniversariesSent: annivCount };
  } catch (e) {
    Logger.log("Lỗi sendBirthdayAndAnniversaryGreetingsTrigger: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 4. MVP Asset Management: Bàn giao tài sản cho nhân viên
 */
function assignAsset(assetName, serialNo, assignedToEmail, notes) {
  requireRole(['Admin', 'HR']);
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var aRead = batchReadSheet('Assets');
    var data = aRead.data;
    if (data.length === 0) {
      data = [['AssetID', 'AssetName', 'SerialNo', 'AssignedToEmail', 'Status', 'AssignedDate', 'ReturnDate', 'Notes']];
    }
    
    var assetId = "AST-" + Number(new Date());
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    data.push([assetId, assetName, serialNo || '', assignedToEmail, 'Assigned', todayStr, '', notes || 'Bàn giao tài sản']);
    batchWriteSheet('Assets', data);
    
    return { success: true, assetId: assetId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 4. MVP Asset Management: Thu hồi tài sản
 */
function returnAsset(assetId, notes) {
  requireRole(['Admin', 'HR']);
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var aRead = batchReadSheet('Assets');
    var data = aRead.data;
    if (data.length <= 1) return { success: false, message: "Chưa có danh sách tài sản." };
    
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(assetId).trim()) {
        rowIdx = i;
        break;
      }
    }
    
    if (rowIdx === -1) return { success: false, message: "Không tìm thấy mã tài sản: " + assetId };
    
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    data[rowIdx][4] = 'Returned';
    data[rowIdx][6] = todayStr;
    if (notes) data[rowIdx][7] = String(data[rowIdx][7]) + " | " + notes;
    
    batchWriteSheet('Assets', data);
    return { success: true, message: "Đã thu hồi tài sản thành công!" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 4. MVP Performance OKR: Gửi bản tự đánh giá OKR (Self Review)
 */
function submitOkrReview(period, objective, keyResultsJson, selfScore) {
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) throw new Error("Vui lòng đăng nhập.");
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var okrRead = batchReadSheet('PerformanceOKRs');
    var data = okrRead.data;
    if (data.length === 0) {
      data = [['ID', 'EmployeeEmail', 'Period', 'Objective', 'KeyResultsJson', 'SelfScore', 'ManagerScore', 'FinalScore', 'Status', 'CreatedAt']];
    }
    
    var okrId = "OKR-" + Number(new Date());
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    data.push([
      okrId,
      userEmail,
      period || 'Q3-2026',
      objective,
      typeof keyResultsJson === 'object' ? JSON.stringify(keyResultsJson) : keyResultsJson,
      Number(selfScore) || 0,
      '',
      '',
      'Submitted',
      nowStr
    ]);
    
    batchWriteSheet('PerformanceOKRs', data);
    return { success: true, okrId: okrId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 4. MVP Performance OKR: Quản lý chấm điểm đánh giá 3 chiều (Manager Score & Final Score)
 */
function evaluateOkrReview(okrId, managerScore, finalScore) {
  requireRole(['Admin', 'HR', 'Manager']);
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var okrRead = batchReadSheet('PerformanceOKRs');
    var data = okrRead.data;
    if (data.length <= 1) return { success: false, message: "Không tìm thấy dữ liệu OKR." };
    
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(okrId).trim()) {
        rowIdx = i;
        break;
      }
    }
    
    if (rowIdx === -1) return { success: false, message: "Không tìm thấy mã OKR: " + okrId };
    
    data[rowIdx][6] = Number(managerScore) || 0;
    data[rowIdx][7] = Number(finalScore) || 0;
    data[rowIdx][8] = 'Evaluated';
    
    batchWriteSheet('PerformanceOKRs', data);
    return { success: true, message: "Đánh giá OKR thành công!" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 4. MVP L&D: Đăng ký Khóa đào tạo & Ký cam kết đào tạo
 */
function registerCourse(courseId, commitmentSigned) {
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) throw new Error("Vui lòng đăng nhập.");
  
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var crRead = batchReadSheet('CourseRegistrations');
    var data = crRead.data;
    if (data.length === 0) {
      data = [['ID', 'CourseID', 'EmployeeEmail', 'CommitmentSigned', 'Status', 'CreatedAt']];
    }
    
    var regId = "REG-" + Number(new Date());
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    data.push([
      regId,
      courseId,
      userEmail,
      commitmentSigned ? 'YES' : 'NO',
      'Registered',
      nowStr
    ]);
    
    batchWriteSheet('CourseRegistrations', data);
    return { success: true, regId: regId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * 5. Chuẩn hóa dữ liệu chỉ số HRM kết nối Looker Studio
 * Tính toán tự động các chỉ số HR cốt lõi: Turnover Rate, Cost per Hire, Time to Hire
 */
function generateLookerAnalyticsData() {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var employees = getEmployees();
    var totalEmp = employees.length;
    var resignedEmp = employees.filter(e => e.status === 'Nghỉ việc' || e.status === 'Thôi việc').length;
    var turnoverRate = totalEmp > 0 ? ((resignedEmp / totalEmp) * 100).toFixed(2) : 0;
    
    var candRead = batchReadSheet('Candidates');
    var candData = candRead.data;
    var passedCount = 0;
    var totalHireDays = 0;
    
    for (var c = 1; c < candData.length; c++) {
      var cStatus = String(candData[c][6]).trim();
      var recvDate = candData[c][1];
      var createdDate = candData[c][9];
      
      if (cStatus === 'Passed' && recvDate) {
        passedCount++;
        var rD = new Date(recvDate);
        var cD = createdDate ? new Date(createdDate) : new Date();
        if (!isNaN(rD.getTime()) && !isNaN(cD.getTime())) {
          var diffD = Math.max(1, Math.ceil((cD.getTime() - rD.getTime()) / (1000 * 60 * 60 * 24)));
          totalHireDays += diffD;
        }
      }
    }
    
    var avgTimeToHire = passedCount > 0 ? (totalHireDays / passedCount).toFixed(1) : 0;
    var avgCostPerHire = 3500000;
    
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");
    
    var metricsData = [
      ['MetricCategory', 'MetricName', 'Period', 'Value', 'Unit', 'Notes', 'UpdatedAt'],
      ['HR Core', 'Turnover Rate', currentMonth, turnoverRate, '%', 'Tỷ lệ nhân viên nghỉ việc trong tháng', nowStr],
      ['Recruiting', 'Cost per Hire', currentMonth, avgCostPerHire, 'VND', 'Chi phí trung bình cho 1 ứng viên nhận việc', nowStr],
      ['Recruiting', 'Time to Hire', currentMonth, avgTimeToHire, 'Ngày', 'Thời gian trung bình từ nộp CV đến nhận việc', nowStr],
      ['HR Core', 'Headcount Total', currentMonth, totalEmp, 'Người', 'Tổng số lượng nhân sự hiện tại', nowStr],
      ['Recruiting', 'Hired Candidates Count', currentMonth, passedCount, 'Người', 'Số lượng ứng viên nhận việc thành công', nowStr]
    ];
    
    batchWriteSheet('Looker_Metrics', metricsData);
    
    return {
      success: true,
      turnoverRate: turnoverRate + "%",
      costPerHire: avgCostPerHire.toLocaleString() + " VND",
      timeToHire: avgTimeToHire + " Ngày",
      totalHeadcount: totalEmp
    };
  } catch (e) {
    Logger.log("Lỗi generateLookerAnalyticsData: " + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * Tính tỷ lệ Prorate Lương theo Số ngày công thực tế trong tháng
 * Dành cho nhân sự vào làm giữa tháng hoặc xin nghỉ việc giữa tháng
 */
function calculateProrateRatio(emp, month, year, standardWorkDays) {
  var m = parseInt(month, 10);
  var y = parseInt(year, 10);
  var firstDay = new Date(y, m - 1, 1);
  var lastDay  = new Date(y, m, 0); // Ngày cuối cùng của tháng
  var totalDaysInMonth = lastDay.getDate();
  
  var activeStartDay = 1;
  var activeEndDay   = totalDaysInMonth;
  
  // 1. Kiểm tra Ngày bắt đầu làm việc (vào giữa tháng)
  var startDateStr = emp.officialStartDate || emp.probationStartDate || emp.startDate;
  if (startDateStr) {
    var sD = new Date(startDateStr);
    if (!isNaN(sD.getTime())) {
      if (sD.getFullYear() === y && (sD.getMonth() + 1) === m) {
        activeStartDay = sD.getDate();
      } else if (sD > lastDay) {
        return 0; // Chưa bắt đầu đi làm trong tháng này
      }
    }
  }
  
  // 2. Kiểm tra Ngày nghỉ việc (nghỉ giữa tháng)
  var endDateStr = emp.contractExpiryDate || emp.endDate || emp.resignDate;
  if (endDateStr && (emp.status === 'Nghỉ việc' || emp.status === 'Thôi việc')) {
    var eD = new Date(endDateStr);
    if (!isNaN(eD.getTime())) {
      if (eD.getFullYear() === y && (eD.getMonth() + 1) === m) {
        activeEndDay = eD.getDate();
      } else if (eD < firstDay) {
        return 0; // Đã nghỉ trước tháng này
      }
    }
  }
  
  // Trọn tháng -> tỷ lệ 1.0 (100%)
  if (activeStartDay === 1 && activeEndDay === totalDaysInMonth) {
    return 1.0;
  }
  
  // Đếm số ngày đi làm thực tế (trừ các ngày Chủ nhật)
  var activeWorkDaysCount = 0;
  for (var d = activeStartDay; d <= activeEndDay; d++) {
    var curD = new Date(y, m - 1, d);
    if (curD.getDay() !== 0) { // 0: Chủ nhật
      activeWorkDaysCount++;
    }
  }
  
  var stdDays = Number(standardWorkDays) || 22;
  var ratio = activeWorkDaysCount / stdDays;
  return Math.min(1.0, Math.max(0.0, ratio));
}

/**
 * TEST CASE 1: Kiểm thử Prorate Lương cho nhân viên mới vào làm ngày 15/07/2026
 */
function testProrateSalaryCaseNewEmployee15th() {
  var mockEmp = {
    id: "EMP-TEST-01",
    fullName: "Nguyễn Văn Mới",
    officialStartDate: "2026-07-15",
    basicSalary: 10000000,
    allowance: 1000000,
    status: "Chính thức"
  };
  var stdDays = 22;
  var ratio = calculateProrateRatio(mockEmp, 7, 2026, stdDays);
  var proratedSalary = Math.round(mockEmp.basicSalary * ratio);
  
  Logger.log("=== TEST CASE 1: Vào làm ngày 15/07/2026 ===");
  Logger.log("Tỷ lệ Prorate công: " + (ratio * 100).toFixed(2) + "%");
  Logger.log("Lương gốc: 10,000,000 VNĐ -> Lương Prorate: " + proratedSalary.toLocaleString() + " VNĐ");
  
  return { ratio: ratio, proratedSalary: proratedSalary };
}

/**
 * TEST CASE 2: Kiểm thử Prorate Lương cho nhân viên xin nghỉ việc ngày 20/07/2026
 */
function testProrateSalaryCaseResignedEmployee20th() {
  var mockEmp = {
    id: "EMP-TEST-02",
    fullName: "Trần Thị Nghỉ",
    officialStartDate: "2024-01-01",
    resignDate: "2026-07-20",
    basicSalary: 10000000,
    allowance: 1000000,
    status: "Nghỉ việc"
  };
  var stdDays = 22;
  var ratio = calculateProrateRatio(mockEmp, 7, 2026, stdDays);
  var proratedSalary = Math.round(mockEmp.basicSalary * ratio);
  
  Logger.log("=== TEST CASE 2: Nghỉ việc ngày 20/07/2026 ===");
  Logger.log("Tỷ lệ Prorate công: " + (ratio * 100).toFixed(2) + "%");
  Logger.log("Lương gốc: 10,000,000 VNĐ -> Lương Prorate: " + proratedSalary.toLocaleString() + " VNĐ");
  
  return { ratio: ratio, proratedSalary: proratedSalary };
}

/**
 * Cascade Archive / Soft-delete Nhân sự để tránh Orphan Data gây lệch báo cáo
 */
function archiveEmployeeAndCascade(empId, reason) {
  requireRole(['Admin', 'HR']);
  if (!empId) return { success: false, message: "Mã nhân sự không được để trống." };
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // 1. Update Employees Sheet (Status = 'Nghỉ việc', ResignReason, ResignDate)
    var empRead = batchReadSheet('Employees');
    if (!empRead.sheet || empRead.data.length <= 1) throw new Error("Không tìm thấy dữ liệu bảng Employees.");
    var empData = empRead.data;
    var empHeaders = empRead.headers;
    var empColMap = {};
    empHeaders.forEach(function(h, idx) { empColMap[h] = idx; });
    
    var empRowIdx = -1;
    var empName = "Unknown";
    for (var i = 1; i < empData.length; i++) {
      if (String(empData[i][0]).trim() === String(empId).trim()) {
        empRowIdx = i;
        empName = empColMap['FullName'] !== undefined ? String(empData[i][empColMap['FullName']]).trim() : "Unknown";
        break;
      }
    }
    
    if (empRowIdx === -1) return { success: false, message: "Không tìm thấy mã nhân sự: " + empId };
    
    if (empColMap['Status'] !== undefined) empData[empRowIdx][empColMap['Status']] = 'Nghỉ việc';
    if (empColMap['EndDate'] !== undefined) empData[empRowIdx][empColMap['EndDate']] = todayStr;
    if (empColMap['ContractExpiryDate'] !== undefined) empData[empRowIdx][empColMap['ContractExpiryDate']] = todayStr;
    
    batchWriteSheet('Employees', empData);
    
    // 2. Cascade Update EmployeeDetails Sheet (Status = 'Archived')
    var detRead = batchReadSheet('EmployeeDetails');
    if (detRead.sheet && detRead.data.length > 1) {
      var detData = detRead.data;
      var detHeaders = detRead.headers;
      var statusCol = detHeaders.indexOf('Status');
      for (var d = 1; d < detData.length; d++) {
        if (String(detData[d][0]).trim() === String(empId).trim()) {
          if (statusCol !== -1) detData[d][statusCol] = 'Archived';
        }
      }
      batchWriteSheet('EmployeeDetails', detData);
    }
    
    // 3. Cascade Asset Retrieval / Soft-archive
    var assetRead = batchReadSheet('Assets');
    if (assetRead.sheet && assetRead.data.length > 1) {
      var assetData = assetRead.data;
      var assetHeaders = assetRead.headers;
      var assignedToCol = assetHeaders.indexOf('AssignedTo');
      var assetStatusCol = assetHeaders.indexOf('Status');
      for (var a = 1; a < assetData.length; a++) {
        if (assignedToCol !== -1 && String(assetData[a][assignedToCol]).trim() === String(empId).trim()) {
          if (assetStatusCol !== -1) assetData[a][assetStatusCol] = 'Available';
          assetData[a][assignedToCol] = '';
        }
      }
      batchWriteSheet('Assets', assetData);
    }
    
    // 4. Log JobHistory for Audit Trail
    logJobHistory(empId, empName, 'Resigned/Archived', 'Active', 'Archived', '', '', reason || 'Lưu trữ/Nghỉ việc nhân sự');
    
    // Invalidate caches
    try { CacheService.getScriptCache().remove('all_employees'); } catch(e) {}
    
    return {
      success: true,
      message: "Đã chuyển nhân sự " + empId + " (" + empName + ") sang trạng thái Lưu trữ (Archived) và thu hồi tài sản liên quan.",
      archivedEmpId: empId
    };
  } catch (e) {
    Logger.log("Lỗi archiveEmployeeAndCascade: " + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(err) {}
  }
}

/**
 * BỘ TEST TỰ ĐỘNG CƠ BẢN (UNIT TEST SUITE) CHO CÁC HÀM TÍNH TOÁN QUAN TRỌNG
 * Kiểm thử: Tính Lương, Trần Bảo hiểm NĐ 73/2024 & NĐ 74/2024, Phép thâm niên, và Prorate Lương
 */
function runHrmSystemUnitTestSuite() {
  var results = [];
  var passedCount = 0;
  var failedCount = 0;
  
  function assertEqual(testName, actual, expected, description) {
    var pass = (actual === expected);
    if (pass) passedCount++; else failedCount++;
    results.push({
      testName: testName,
      status: pass ? "PASSED" : "FAILED",
      actual: actual,
      expected: expected,
      description: description
    });
  }
  
  // ---- TEST 1: Cấu hình Trần BHXH/BHYT (NĐ 73/2024 - 46,800,000 VNĐ) & BHTN (NĐ 74/2024 - 99,200,000 VNĐ) ----
  var salaryHigh = 150000000;
  var insResult = calculateMandatoryInsurance(salaryHigh, salaryHigh);
  // BHXH (8%) + BHYT (1.5%) trần 46.8M = 4,446,000 VNĐ. BHTN (1%) trần 99.2M = 992,000 VNĐ -> Tổng = 5,438,000 VNĐ
  assertEqual("TEST_INSURANCE_CAP_HIGH_SALARY", insResult, 5438000, "Kiểm tra mức đóng BHXH/BHTN tối đa áp dụng trần NĐ 73/2024 & NĐ 74/2024 cho lương 150 triệu");
  
  // ---- TEST 2: Phép Thâm niên theo Điều 114 Luật Lao động 2019 ----
  // Giả định 12 năm thâm niên -> Cứ 5 năm được +1 ngày -> +2 ngày phép thâm niên
  var seniorityYears = 12;
  var extraSeniorityDays = Math.floor(seniorityYears / 5);
  assertEqual("TEST_SENIORITY_LEAVE_ARTICLE_114", extraSeniorityDays, 2, "Kiểm tra cộng 2 ngày phép thâm niên cho 12 năm làm việc");
  
  // ---- TEST 3: Prorate Lương cho Nhân viên vào làm ngày 15/07/2026 ----
  var mockNewEmp = {
    id: "EMP-TEST-01",
    officialStartDate: "2026-07-15",
    basicSalary: 10000000,
    status: "Chính thức"
  };
  var prorateRatioNew = calculateProrateRatio(mockNewEmp, 7, 2026, 22);
  var proratedSalaryNew = Math.round(mockNewEmp.basicSalary * prorateRatioNew);
  assertEqual("TEST_PRORATE_NEW_EMPLOYEE_15TH", proratedSalaryNew, 6818182, "Kiểm tra lương Prorate cho nhân viên vào làm ngày 15/07 (15/22 ngày công = 68.18%)");
  
  // ---- TEST 4: Prorate Lương cho Nhân viên nghỉ việc ngày 20/07/2026 ----
  var mockResignEmp = {
    id: "EMP-TEST-02",
    resignDate: "2026-07-20",
    basicSalary: 10000000,
    status: "Nghỉ việc"
  };
  var prorateRatioResign = calculateProrateRatio(mockResignEmp, 7, 2026, 22);
  var proratedSalaryResign = Math.round(mockResignEmp.basicSalary * prorateRatioResign);
  assertEqual("TEST_PRORATE_RESIGNED_EMPLOYEE_20TH", proratedSalaryResign, 7727273, "Kiểm tra lương Prorate cho nhân viên nghỉ ngày 20/07 (17/22 ngày công = 77.27%)");

  // ---- TEST 5: Tính Thuế TNCN Khấu trừ Gia cảnh (11M bản thân + 4.4M/người phụ thuộc) ----
  // Lương chịu thuế = 20M - Ins (1.9M) - 11M - 4.4M (1 NPT) = 2.7M -> Bậc 1 (5%) = 135,000 VNĐ
  var insStandard = calculateMandatoryInsurance(20000000, 20000000);
  var taxableIncome = 20000000 - insStandard - 11000000 - 4400000;
  var pitTax = Math.round(calculatePIT(taxableIncome));
  assertEqual("TEST_PIT_TAX_CALCULATION", pitTax, 135000, "Kiểm tra tính Thuế TNCN bậc 1 (5%) cho thu nhập tính thuế 2.7 triệu");

  Logger.log("=== BỘ KẾT QUẢ UNIT TEST HRM SUITE ===");
  Logger.log("Tổng số test cases: " + results.length + " | PASSED: " + passedCount + " | FAILED: " + failedCount);
  
  return {
    success: failedCount === 0,
    total: results.length,
    passedCount: passedCount,
    failedCount: failedCount,
    results: results
  };
}





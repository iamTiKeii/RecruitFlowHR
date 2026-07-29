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
 * Secure SHA-256 Hashing helper
 */
function hashPassword(password) {
    var salt = "MY_FIXED_SALT_123";
    var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt, Utilities.Charset.UTF_8);
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

/**
 * Authentication helper matching Users sheet with CacheService caching (10 minutes)
 */
/**
 * Authentication helper matching Users sheet with CacheService caching (10 minutes)
 */
function authenticate(usernameOrEmail, password) {
    var inputHash = hashPassword(password);
    var cleanEmail = String(usernameOrEmail).toLowerCase().trim();
    var cacheKey = 'auth_' + cleanEmail + '_' + inputHash;
    var cache = CacheService.getScriptCache();

    try {
        var cachedVal = cache.get(cacheKey);
        if (cachedVal) {
            return JSON.parse(cachedVal);
        }
    } catch (e) {
        // Ignore cache lookup issues
    }

    var ss = getSpreadsheet();
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
        initializeDatabaseV2();
        usersSheet = ss.getSheetByName('Users');
    }

    var usersData = usersSheet.getDataRange().getValues();
    for (var i = 1; i < usersData.length; i++) {
        var sheetEmail = usersData[i][0] ? String(usersData[i][0]).trim() : "";
        var sheetHash = usersData[i][1] ? String(usersData[i][1]).trim() : "";
        if (sheetEmail.toLowerCase() === cleanEmail) {
            if (sheetHash === inputHash) {
                var userObj = {
                    success: true,
                    user: {
                        username: usersData[i][0] ? String(usersData[i][0]).trim() : "",
                        email: usersData[i][0] ? String(usersData[i][0]).trim() : "",
                        fullName: usersData[i][2] ? String(usersData[i][2]).trim() : "",
                        role: usersData[i][3] ? String(usersData[i][3]).trim() : ""
                    }
                };
                try {
                    cache.put(cacheKey, JSON.stringify(userObj), 600); // cache for 10 minutes (600s)
                } catch (e) { }
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
 * Password change helper (clears cache keys)
 */
function changePassword(usernameOrEmail, oldPassword, newPassword) {
    var ss = getSpreadsheet();
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
        initializeDatabaseV2();
        usersSheet = ss.getSheetByName('Users');
    }

    var usersData = usersSheet.getDataRange().getValues();
    var userRow = -1;
    var savedHash = "";
    var targetEmail = String(usernameOrEmail).toLowerCase().trim();

    for (var i = 1; i < usersData.length; i++) {
        var sheetEmail = usersData[i][0] ? String(usersData[i][0]).trim() : "";
        if (sheetEmail.toLowerCase() === targetEmail) {
            userRow = i + 1;
            savedHash = usersData[i][1] ? String(usersData[i][1]).trim() : "";
            break;
        }
    }

    if (userRow === -1) {
        return { success: false, message: "Không tìm thấy người dùng." };
    }

    if (savedHash !== hashPassword(oldPassword)) {
        return { success: false, message: "Mật khẩu cũ không chính xác." };
    }

    var newHash = hashPassword(newPassword);
    usersSheet.getRange(userRow, 2).setValue(newHash);
    // Clear auth CacheService
    var cache = CacheService.getScriptCache();
    try {
        cache.remove('auth_' + targetEmail + '_' + savedHash);
        cache.remove('auth_' + targetEmail + '_' + newHash);
    } catch (e) { }

    return { success: true };
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
            'Users': ['Email', 'PasswordHash', 'FullName', 'Role', 'Department'],
            'Categories': ['ID', 'CategoryType', 'Code', 'Name', 'Value', 'Description'],
            'Employees': [
                'ID', 'FullName', 'Gender', 'Email', 'Phone', 'Department', 'Position',
                'BasicSalary', 'ProbationSalary', 'Allowances', 'ProbationStartDate',
                'OfficialStartDate', 'ContractExpiryDate', 'Status'
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
                'ProbationSalary', 'CurrencyUnit', 'AssignedInterviewer', 'MeetLink', 'OfferData', 'EvaluationHistory', 'EmailContent'
            ],
            'Interviews': ['ID', 'CandidateID', 'InterviewDate', 'Location', 'Note'],
            'Forms': ['FormID', 'FormName', 'PublishedUrl', 'EditUrl', 'ResponseSheetName', 'CreatedAt'],
            'Attendance': ['ID', 'EmployeeID', 'Date', 'CheckIn', 'CheckOut', 'IPAddress', 'OT_Hours', 'Status'],
            'LeaveRequests': ['ID', 'EmployeeID', 'LeaveType', 'StartDate', 'EndDate', 'Reason', 'Approver', 'Status'],
            'Payroll': ['ID', 'Month', 'Year', 'EmployeeID', 'TotalSalary', 'Insurance', 'Tax', 'NetSalary', 'SentDate'],
            'EmployeeDetails': [
                'EmployeeID', 'FullName', 'Gender', 'BirthPlace', 'CurrentAddress', 'RegisterAddress',
                'IdentityCardNumber', 'IdentityCardDate', 'IdentityCardPlace', 'AcademicLevel',
                'Specialization', 'GraduationInstitution', 'YouthUnionDate', 'CommunistPartyDateStatus', 'Docs', 'DateOfBirth', 'Department'
            ],
            'SalaryHistory': ['ID', 'EmployeeID', 'NewSalary', 'ChangeDate', 'FileBase64', 'Notes']
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
function getConfig() {
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
    return config;
}

/**
 * Save configuration configurations (safe checks)
 */
function saveConfig(configData) {
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

    return { success: true };
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
        receivedDate: data[rowIndex - 1][1] ? String(data[rowIndex - 1][1]).trim() : "",
        senderName: data[rowIndex - 1][2] ? String(data[rowIndex - 1][2]).trim() : "",
        senderEmail: data[rowIndex - 1][3] ? String(data[rowIndex - 1][3]).trim() : "",
        subject: data[rowIndex - 1][4] ? String(data[rowIndex - 1][4]).trim() : "",
        cvLink: data[rowIndex - 1][5] ? String(data[rowIndex - 1][5]).trim() : "",
        status: status,
        dateOfBirth: data[rowIndex - 1][11] ? String(data[rowIndex - 1][11]).trim() : "",
        phoneNumber: data[rowIndex - 1][12] ? String(data[rowIndex - 1][12]).trim() : ""
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

    var getValue = function (row, colName, fallbackIdx) {
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

    candidates.sort(function (a, b) {
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
    var beforeNames = beforeSheets.map(function (s) { return s.getName(); });

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
        receivedDate: data[rowIndex - 1][1] ? String(data[rowIndex - 1][1]).trim() : "",
        senderName: data[rowIndex - 1][2] ? String(data[rowIndex - 1][2]).trim() : "",
        senderEmail: data[rowIndex - 1][3] ? String(data[rowIndex - 1][3]).trim() : "",
        subject: data[rowIndex - 1][4] ? String(data[rowIndex - 1][4]).trim() : "",
        cvLink: data[rowIndex - 1][5] ? String(data[rowIndex - 1][5]).trim() : "",
        status: "Suitable",
        dateOfBirth: data[rowIndex - 1][11] ? String(data[rowIndex - 1][11]).trim() : "",
        phoneNumber: data[rowIndex - 1][12] ? String(data[rowIndex - 1][12]).trim() : "",
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
                var question = headers[c] ? String(headers[c]).trim() : "Question " + (c + 1);
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
function deleteCandidate(candidateId) {
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
            } catch (e) {
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
    } catch (e) {
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

    var getValue = function (row, colName, fallbackIdx) {
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
        } catch (e) {
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
                return {
                    email: email,
                    role: data[i][3] ? String(data[i][3]).trim() : "",
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

    return { email: email, role: "", fullName: "" };
}

/**
 * Helper to fetch a user's role by email.
 */
function getUserRoleByEmail(email) {
    if (!email) return "";
    email = email.toLowerCase().trim();
    var ss = getSpreadsheet();
    var usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
        var data = usersSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            var sheetEmail = data[i][0] ? String(data[i][0]).toLowerCase().trim() : "";
            if (sheetEmail === email) {
                return data[i][3] ? String(data[i][3]).trim() : "";
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

    return "";
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
    } catch (sessErr) {
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
function getInterviewerCandidates(clientEmail) {
    var activeEmail = clientEmail || Session.getActiveUser().getEmail();
    if (!activeEmail) return [];
    activeEmail = activeEmail.toLowerCase().trim();

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
function getCandidateCVUrlOrBase64(candidateId, clientEmail) {
    var activeEmail = clientEmail || Session.getActiveUser().getEmail();
    if (!activeEmail) throw new Error("Yêu cầu xác thực tài khoản.");
    activeEmail = activeEmail.toLowerCase().trim();

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
function updateInterviewResult(candidateId, result, offerData, clientEmail) {
    var activeEmail = clientEmail || Session.getActiveUser().getEmail();
    if (!activeEmail) throw new Error("Yêu cầu đăng nhập.");
    activeEmail = activeEmail.toLowerCase().trim();

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

    var headers = data[0].map(function (h) { return String(h).trim(); });
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
            var headers = candData[0].map(function (h) { return String(h).trim(); });
            var idCol = headers.indexOf('ID');
            var nameCol = headers.indexOf('SenderName');
            var emailCol = headers.indexOf('SenderEmail');
            var phoneCol = headers.indexOf('PhoneNumber');
            var statusCol = headers.indexOf('Status');
            var grossCol = headers.indexOf('OfferGross');
            var netCol = headers.indexOf('OfferNet');
            var probCol = headers.indexOf('ProbationSalary');

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
 * Get list of all employees
 */
function getEmployees() {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Employees');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var headers = data[0].map(function (h) { return String(h).trim(); });
    var colMap = {};
    headers.forEach(function (h, idx) {
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
            rawAllowances.split(',').forEach(function (code) {
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
    return employees;
}

/**
 * THÊM NHÂN SỰ MỚI: Đồng thời chèn dữ liệu đồng bộ sang Employees và EmployeeDetails
 */
function addEmployee(emp) {
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

    // Lấy ngày sinh (DOB) từ Candidate nếu có nhập từ phần tuyển dụng
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
    setDetVal('Avatar', emp.avatar || "");
    setDetVal('Docs', "[]");

    detSheet.appendRow(detRow);
    SpreadsheetApp.flush();

    syncEmployeeToUsers(emp.email, emp.fullName, emp.position);

    if (emp.candidateId) {
        updateCandidateStatus(emp.candidateId, 'Hired');
    }

    return { success: true, employeeId: nextId };
}

/**
 * Update employee details
 */
function updateEmployee(emp) {
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
        emp.allowances.split(',').forEach(function (code) {
            var trimmed = code.trim();
            if (allowanceValuesMap[trimmed]) {
                totalAllowance += allowanceValuesMap[trimmed];
            }
        });
    }
    emp.allowance = totalAllowance;

    var headers = data[0].map(function (h) { return String(h).trim(); });
    var colMap = {};
    headers.forEach(function (h, idx) {
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

    return { success: true };
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

        var headers = data[0].map(function (h) { return String(h).trim(); });
        var colMap = {};
        headers.forEach(function (h, idx) { colMap[h] = idx; });

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
                        var candHeaders = candData[0].map(function (h) { return String(h).trim(); });
                        var candColMap = {};
                        candHeaders.forEach(function (h, idx) { candColMap[h] = idx; });
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
 * Fetch attendance logs based on RBAC filters
 */
function getAttendanceLogs(email, role) {
    var ss = getSpreadsheet();
    var attSheet = ss.getSheetByName('Attendance');
    var empSheet = ss.getSheetByName('Employees');
    if (!attSheet || !empSheet) return [];

    var empData = empSheet.getDataRange().getValues();
    var empIdToNameMap = {};
    var empIdToEmailMap = {};
    var empIdToDeptMap = {};
    var targetEmployeeId = "";
    var userEmail = String(email || "").toLowerCase().trim();

    for (var i = 1; i < empData.length; i++) {
        var id = String(empData[i][0]).trim();
        var name = String(empData[i][1]).trim();
        var empEmail = String(empData[i][2]).toLowerCase().trim();
        var dept = String(empData[i][4]).trim();
        empIdToNameMap[id] = name;
        empIdToEmailMap[id] = empEmail;
        empIdToDeptMap[id] = dept;

        if (empEmail === userEmail) {
            targetEmployeeId = id;
        }
    }

    var logs = [];
    var data = attSheet.getDataRange().getValues();
    for (var j = 1; j < data.length; j++) {
        var empId = String(data[j][1]).trim();
        var logEmail = empIdToEmailMap[empId] || "";
        var logDept = empIdToDeptMap[empId] || "";

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
function submitLeaveRequest(leave) {
    var ss = getSpreadsheet();
    var lrSheet = ss.getSheetByName('LeaveRequests');
    var empSheet = ss.getSheetByName('Employees');
    if (!lrSheet || !empSheet) throw new Error("Database sheets not initialized");

    var empData = empSheet.getDataRange().getValues();
    var empId = "";
    var empName = "";
    var empDept = "";
    var userEmail = String(leave.employeeEmail || "").toLowerCase().trim();
    for (var i = 1; i < empData.length; i++) {
        if (String(empData[i][2]).toLowerCase().trim() === userEmail) {
            empId = String(empData[i][0]).trim();
            empName = String(empData[i][1]).trim();
            empDept = String(empData[i][4]).trim();
            break;
        }
    }

    if (!empId) throw new Error("Employee email not found.");

    var lrData = lrSheet.getDataRange().getValues();
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
    for (var i = 1; i < empData.length; i++) {
        var dept = String(empData[i][4]).trim();
        var pos = String(empData[i][5]).toLowerCase().trim();
        if (dept === empDept && (pos.indexOf("manager") !== -1 || pos.indexOf("trưởng phòng") !== -1 || pos.indexOf("lead") !== -1)) {
            managerEmail = String(empData[i][2]).trim();
            break;
        }
    }

    if (!managerEmail) {
        managerEmail = "hr@recruitflow.com";
    }

    lrSheet.appendRow([
        nextId,
        empId,
        leave.leaveType || "Nghỉ phép năm",
        leave.startDate || "",
        leave.endDate || "",
        leave.reason || "",
        managerEmail,
        "Pending"
    ]);

    try {
        var emailBody = "<div style='font-family: sans-serif; padding: 20px; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px;'>" +
            "<h2>YÊU CẦU DUYỆT NGHỈ PHÉP</h2>" +
            "<p>Kính gửi Quản lý,</p>" +
            "<p>Nhân viên <strong>" + empName + "</strong> (Phòng ban: " + empDept + ") vừa gửi yêu cầu nghỉ phép cần bạn phê duyệt:</p>" +
            "<table style='width: 100%; border-collapse: collapse; margin: 15px 0;'>" +
            "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Loại phép:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + (leave.leaveType || "Nghỉ phép năm") + "</td></tr>" +
            "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Thời gian:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>Từ " + leave.startDate + " đến " + leave.endDate + "</td></tr>" +
            "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><strong>Lý do:</strong></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + (leave.reason || "Không có") + "</td></tr>" +
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

    return { success: true };
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
        var email = String(empData[i][2]).toLowerCase().trim();
        var dept = String(empData[i][4]).trim();
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
                var e = String(empData[k][2]).toLowerCase().trim();
                if (e === cleanUserEmail) {
                    managerDept = String(empData[k][4]).trim();
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
 * Approve or Reject leave requests
 */
function approveLeaveRequest(requestId, status, approverEmail) {
    var ss = getSpreadsheet();
    var lrSheet = ss.getSheetByName('LeaveRequests');
    var empSheet = ss.getSheetByName('Employees');
    if (!lrSheet || !empSheet) throw new Error("Sheets not found");

    var data = lrSheet.getDataRange().getValues();
    var rowIndex = -1;
    var employeeId = "";
    for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === requestId) {
            rowIndex = i + 1;
            employeeId = String(data[i][1]).trim();
            break;
        }
    }

    if (rowIndex === -1) throw new Error("Leave request not found");

    lrSheet.getRange(rowIndex, 8).setValue(status);

    var empData = empSheet.getDataRange().getValues();
    var empEmail = "";
    var empName = "";
    for (var j = 1; j < empData.length; j++) {
        if (String(empData[j][0]).trim() === employeeId) {
            empEmail = String(empData[j][2]).trim();
            empName = String(empData[j][1]).trim();
            break;
        }
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
        } catch (err) {
            Logger.log("Failed to send leave result email: " + err.message);
        }
    }

    return { success: true };
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
 * Perform monthly payroll calculations
 */
function calculatePayroll(month, year) {
    var ss = getSpreadsheet();
    var empSheet = ss.getSheetByName('Employees');
    var attSheet = ss.getSheetByName('Attendance');
    var paySheet = ss.getSheetByName('Payroll');
    if (!empSheet || !attSheet || !paySheet) throw new Error("Database sheets not initialized");

    var configs = getConfig();
    var hourlyRate = Number(configs.ot_hourly_rate || 150000);

    var employees = getEmployees();
    var activeEmployees = employees.filter(e => e.status === "Đang làm");

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

    var otMap = {};
    attendanceLogs.forEach(log => {
        if (!otMap[log.employeeId]) otMap[log.employeeId] = 0;
        otMap[log.employeeId] += log.otHours;
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
        var emp = activeEmployees[e];
        var otHours = otMap[emp.id] || 0;

        var basicSalary = emp.basicSalary || 0;
        var allowance = emp.allowance || 0;
        var otSalary = otHours * hourlyRate;

        var gross = basicSalary + allowance + otSalary;
        var insurance = Math.round(basicSalary * 0.105);
        var taxable = gross - insurance - 11000000;
        var tax = Math.round(calculatePIT(taxable));
        var net = gross - insurance - tax;

        var key = emp.id + "_" + month + "_" + year;
        var rowIdx = payRowMap[key];

        if (rowIdx) {
            paySheet.getRange(rowIdx, 5).setValue(gross);
            paySheet.getRange(rowIdx, 6).setValue(insurance);
            paySheet.getRange(rowIdx, 7).setValue(tax);
            paySheet.getRange(rowIdx, 8).setValue(net);
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
                ""
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
 * Fetch payroll logs for personal / management views
 */
function getPayslips(email, role) {
    var ss = getSpreadsheet();
    var paySheet = ss.getSheetByName('Payroll');
    var empSheet = ss.getSheetByName('Employees');
    if (!paySheet || !empSheet) return [];

    var empData = empSheet.getDataRange().getValues();
    var empIdToNameMap = {};
    var empIdToEmailMap = {};
    var empIdToDeptMap = {};
    var targetEmployeeId = "";
    var cleanUserEmail = String(email || "").toLowerCase().trim();

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
 * Get all Master Data categories
 */
function getCategories() {
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
        var empHeaders = empData[0].map(function (h) { return String(h).trim(); });
        var empColMap = {};
        empHeaders.forEach(function (h, idx) { empColMap[h] = idx; });

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
        var detHeaders = detData[0].map(function (h) { return String(h).trim(); });
        var detColMap = {};
        detHeaders.forEach(function (h, idx) { detColMap[h] = idx; });

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
            var payrollHeaders = payrollData[0].map(function (h) { return String(h).trim(); });
            var pyEmpIdx = payrollHeaders.indexOf('EmployeeID');
            for (var i = 1; i < payrollData.length; i++) {
                if (pyEmpIdx !== -1 && String(payrollData[i][pyEmpIdx]).trim() === employeeId) {
                    var pr = {};
                    payrollHeaders.forEach(function (h, idx) {
                        pr[h.charAt(0).toLowerCase() + h.slice(1)] = payrollData[i][idx];
                    });
                    pr.id = String(payrollData[i][payrollHeaders.indexOf('ID') !== -1 ? payrollHeaders.indexOf('ID') : 0]).trim();
                    pr.month = payrollData[i][payrollHeaders.indexOf('Month') !== -1 ? payrollHeaders.indexOf('Month') : 1] || '';
                    pr.year = payrollData[i][payrollHeaders.indexOf('Year') !== -1 ? payrollHeaders.indexOf('Year') : 2] || '';
                    pr.totalSalary = Number(payrollData[i][payrollHeaders.indexOf('TotalSalary') !== -1 ? payrollHeaders.indexOf('TotalSalary') : 4]) || 0;
                    pr.insurance = Number(payrollData[i][payrollHeaders.indexOf('Insurance') !== -1 ? payrollHeaders.indexOf('Insurance') : 5]) || 0;
                    pr.tax = Number(payrollData[i][payrollHeaders.indexOf('Tax') !== -1 ? payrollHeaders.indexOf('Tax') : 6]) || 0;
                    pr.netSalary = Number(payrollData[i][payrollHeaders.indexOf('NetSalary') !== -1 ? payrollHeaders.indexOf('NetSalary') : 7]) || 0;
                    pr.sentDate = formatDate(payrollData[i][payrollHeaders.indexOf('SentDate') !== -1 ? payrollHeaders.indexOf('SentDate') : 8]);
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
            var candHeaders = candData[0].map(function (h) { return String(h).trim(); });
            var cEmailIdx = candHeaders.indexOf('SenderEmail');
            var cIdIdx = candHeaders.indexOf('ID');
            var cNameIdx = candHeaders.indexOf('SenderName');
            var cCVIdx = candHeaders.indexOf('CV_Link');
            var cStatusIdx = candHeaders.indexOf('Status');
            var cInterIdx = candHeaders.indexOf('InterviewInfo');
            var cEvalIdx = candHeaders.indexOf('EvaluationHistory');
            var cDOBIdx = candHeaders.indexOf('DateOfBirth');
            var cPhoneIdx = candHeaders.indexOf('PhoneNumber');
            var cOfferGross = candHeaders.indexOf('OfferGross');
            var cOfferNet = candHeaders.indexOf('OfferNet');
            var cOfferStart = candHeaders.indexOf('OfferStartDate');
            var cSubjectIdx = candHeaders.indexOf('Subject');
            var cRecvIdx = candHeaders.indexOf('ReceivedDate');

            for (var i = 1; i < candData.length; i++) {
                var rowEmail = cEmailIdx !== -1 ? String(candData[i][cEmailIdx]).trim().toLowerCase() : '';
                if (rowEmail && rowEmail === empEmail.toLowerCase()) {
                    candidateObj = {
                        id: cIdIdx !== -1 ? String(candData[i][cIdIdx]).trim() : '',
                        senderName: cNameIdx !== -1 ? String(candData[i][cNameIdx]).trim() : '',
                        senderEmail: empEmail,
                        subject: cSubjectIdx !== -1 ? String(candData[i][cSubjectIdx]).trim() : '',
                        cvLink: cCVIdx !== -1 ? String(candData[i][cCVIdx]).trim() : '',
                        status: cStatusIdx !== -1 ? String(candData[i][cStatusIdx]).trim() : '',
                        interviewInfo: cInterIdx !== -1 ? String(candData[i][cInterIdx]).trim() : '',
                        evaluationHistory: cEvalIdx !== -1 ? String(candData[i][cEvalIdx]).trim() : '',
                        dateOfBirth: cDOBIdx !== -1 ? formatDate(candData[i][cDOBIdx]) : '',
                        phone: cPhoneIdx !== -1 ? String(candData[i][cPhoneIdx]).trim() : '',
                        offerGross: cOfferGross !== -1 ? Number(candData[i][cOfferGross]) || 0 : 0,
                        offerNet: cOfferNet !== -1 ? Number(candData[i][cOfferNet]) || 0 : 0,
                        offerStartDate: cOfferStart !== -1 ? formatDate(candData[i][cOfferStart]) : '',
                        receivedDate: cRecvIdx !== -1 ? formatDate(candData[i][cRecvIdx]) : ''
                    };
                    break;
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
            candidate: candidateObj
        };

    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

/**
 * LUU THAY DOI CHI TIET: Cap nhat dong bo ca 2 bang Employees va EmployeeDetails
 * Tu khac phuc: Neu EmployeeDetails chua co dong thi tu tao truoc khi luu
 */
function saveEmployeeDetail(details) {
    try {
        var ss = getSpreadsheet();
        var employeeId = details.employeeId || details.id || '';
        if (!employeeId) throw new Error('ID nhan su khong hop le. Kiem tra details.employeeId hoac details.id.');

        // ---- 1. CAP NHAT BANG EMPLOYEES ----
        var empSheet = ss.getSheetByName('Employees');
        if (empSheet) {
            var empData = empSheet.getDataRange().getValues();
            var empHeaders = empData[0].map(function (h) { return String(h).trim(); });
            var empColMap = {};
            empHeaders.forEach(function (h, idx) { empColMap[h] = idx; });

            var empRowIdx = -1;
            for (var i = 1; i < empData.length; i++) {
                if (String(empData[i][0]).trim() === employeeId) { empRowIdx = i + 1; break; }
            }

            if (empRowIdx !== -1) {
                function updateEmpCell(colName, value) {
                    var idx = empColMap[colName];
                    if (idx !== undefined && value !== undefined && value !== null) {
                        empSheet.getRange(empRowIdx, idx + 1).setValue(value);
                    }
                }
                updateEmpCell('FullName', details.fullName);
                updateEmpCell('Gender', details.gender);
                updateEmpCell('Email', details.email);
                updateEmpCell('Phone', details.phone);
                updateEmpCell('Department', details.department);
                updateEmpCell('Position', details.position);
                updateEmpCell('BasicSalary', details.basicSalary ? Number(details.basicSalary) : 0);
                updateEmpCell('ProbationSalary', details.probationSalary ? Number(details.probationSalary) : 0);
                updateEmpCell('Allowances', details.allowances);
                updateEmpCell('ProbationStartDate', details.probationStartDate);
                updateEmpCell('OfficialStartDate', details.officialStartDate);
                updateEmpCell('ContractExpiryDate', details.contractExpiryDate);
                updateEmpCell('Status', details.status);
                // Legacy compat + ContractType
                updateEmpCell('ContractType', details.contractType);
                updateEmpCell('StartDate', details.officialStartDate || details.probationStartDate);
                updateEmpCell('EndDate', details.contractExpiryDate);
            }
        }

        // ---- 2. CAP NHAT / TU KHAC PHUC BANG EMPLOYEEDETAILS ----
        var detSheet = ss.getSheetByName('EmployeeDetails');
        if (!detSheet) {
            initializeDatabaseV2();
            detSheet = ss.getSheetByName('EmployeeDetails');
        }
        if (detSheet) {
            var detData = detSheet.getDataRange().getValues();
            var detHeaders = detData[0].map(function (h) { return String(h).trim(); });
            var detColMap = {};
            detHeaders.forEach(function (h, idx) { detColMap[h] = idx; });

            var detRowIdx = -1;
            for (var i = 1; i < detData.length; i++) {
                if (String(detData[i][0]).trim() === employeeId) { detRowIdx = i + 1; break; }
            }

            // Self-healing: tu tao dong neu thieu
            if (detRowIdx === -1) {
                var healRow = new Array(detHeaders.length).fill('');
                if (detColMap['EmployeeID'] !== undefined) healRow[detColMap['EmployeeID']] = employeeId;
                if (detColMap['FullName'] !== undefined) healRow[detColMap['FullName']] = details.fullName || '';
                if (detColMap['Department'] !== undefined) healRow[detColMap['Department']] = details.department || '';
                if (detColMap['Gender'] !== undefined) healRow[detColMap['Gender']] = details.gender || 'Nam';
                if (detColMap['Docs'] !== undefined) healRow[detColMap['Docs']] = '[]';
                detSheet.appendRow(healRow);
                SpreadsheetApp.flush();
                detData = detSheet.getDataRange().getValues();
                for (var i = 1; i < detData.length; i++) {
                    if (String(detData[i][0]).trim() === employeeId) { detRowIdx = i + 1; break; }
                }
            }

            if (detRowIdx !== -1) {
                function updateDetCell(colName, value) {
                    var idx = detColMap[colName];
                    if (idx !== undefined && value !== undefined && value !== null) {
                        detSheet.getRange(detRowIdx, idx + 1).setValue(value);
                    }
                }
                updateDetCell('FullName', details.fullName);
                updateDetCell('Department', details.department);
                updateDetCell('Avatar', details.avatar);
                updateDetCell('Gender', details.gender);
                updateDetCell('BirthPlace', details.birthPlace);
                updateDetCell('CurrentAddress', details.currentAddress);
                updateDetCell('RegisterAddress', details.registerAddress);
                updateDetCell('IdentityCardNumber', details.idCardNumber);
                updateDetCell('IdentityCardDate', details.idCardDate);
                updateDetCell('IdentityCardPlace', details.idCardPlace);
                updateDetCell('AcademicLevel', details.academicLevel);
                updateDetCell('Specialization', details.specialization);
                updateDetCell('GraduationInstitution', details.graduationInstitution);
                updateDetCell('YouthUnionDate', details.youthUnionDate);
                updateDetCell('CommunistPartyDateStatus', details.communistPartyDateStatus);
                updateDetCell('Docs', details.docs);
                updateDetCell('DateOfBirth', details.dob || details.dateOfBirth);
            }
        }

        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
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
    } catch (e) {
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

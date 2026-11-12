/**
 * Print & Copy Backend API (V24.0 - Full Production Edition)
 * รองรับ Vercel Frontend, In-app Preview, Telegram Alert, Chat System & Quota Tracking
 */

const CONFIG = {
  SHEET_ID: "10CjyRhOnMfL8iy-HdDbxYZVxbaHOMtoipkpQ9MsoRb8",
  FOLDER_F1_ID: "1KGA9th3OGYNShJ-Flk9-aj6-MiZxZaxL",
  FOLDER_F2_ID: "19LiwQr8TwD1AQvrLTqhTwweTPCjaNuFj",
  FOLDER_F3_ID: "14VJYmne_KAZ1ijIUhwtYqRtM2swZSxxa",
  TIMEZONE: "Asia/Bangkok",

  RENAME_RULES: [
    { from: /\boouserround\b/g, to: "00oouserround" },
    { from: /\b12achiraya_na\b/g, to: "13achiraya_na" },
    { from: /\b12suphapron_si\b/g, to: "07suphapron_si" }
  ],

  SHEETS: {
    raw_f1: "raw_f1", raw_f2: "raw_f2", raw_f3: "raw_f3",
    agg_f1: "agg_f1", agg_f2: "agg_f2", agg_f3: "agg_f3",
    agg_all: "agg_all", staff_map: "Staff_Map", agg_staff: "agg_staff",
    users: "App_Users", settings: "App_Settings", groups: "App_Groups",
    print_req: "Print_Requests", chats: "App_Chats"
  },
  
  DO_ARCHIVE: true
};

// ==========================================
// 1. API ROUTER (ตัวรับคำสั่งจาก Vercel Web App)
// ==========================================
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var data = params.data || {};
    
    // Auth & User Management
    if (action === 'loginUser') return apiResponse(loginUser(data.mobile, data.pin));
    if (action === 'registerUser') return apiResponse(registerUser(data));
    if (action === 'requestNewPin') return apiResponse(requestNewPin(data.mobile));
    if (action === 'adminCreateUser') return apiResponse(adminCreateUser(data));
    if (action === 'approveUser') return apiResponse(approveUser(data.mobile));
    if (action === 'suspendUser') return apiResponse(suspendUser(data.mobile));
    if (action === 'deleteUser') return apiResponse(deleteUser(data.mobile));
    if (action === 'editUser') return apiResponse(editUser(data.mobile, data.updatedData));
    
    // Staff Verification (verify.html)
    if (action === 'getStaffByPrinterCode') return apiResponse(getStaffByPrinterCode(data.printerCode));
    if (action === 'updateStaffData') return apiResponse(updateStaffData(data));
    
    // Self-Service Profile Edit
    if (action === 'requestProfileEdit') return apiResponse(requestProfileEdit(data.mobile, data.newData));
    if (action === 'approveProfileEdit') return apiResponse(approveProfileEdit(data.mobile, data.decision));
    
    // Print Requests
    if (action === 'submitPrintRequest') return apiResponse(submitPrintRequest(data));
    if (action === 'updatePrintReqStatus') return apiResponse(updatePrintReqStatus(data.tsStr, data.newStatus));
    if (action === 'deletePrintReq') return apiResponse(deletePrintReq(data.tsStr));
    
    // Chat System
    if (action === 'sendChatMessage') return apiResponse(sendChatMessage(data));
    if (action === 'editChatMessage') return apiResponse(editChatMessage(data));
    if (action === 'deleteChatMessage') return apiResponse(deleteChatMessage(data));
    
    // Data & Dashboard
    if (action === 'getDashboardData') return apiResponse({ success: true, data: getDashboardData(data.mobile) });
    if (action === 'getGroupOptions') return apiResponse({ success: true, data: getGroupOptions() });
    if (action === 'getAdminData') return apiResponse({ success: true, data: getAdminData() });
    
    // Settings & ETL
    if (action === 'saveSettings') return apiResponse(saveSettings(data.costBw, data.costColor, data.emails, data.telegramToken, data.telegramChatId));
    if (action === 'saveGroupConfig') return apiResponse(saveGroupConfig(data.groups));
    if (action === 'saveGroupQuotas') return apiResponse(saveGroupQuotas(data.quotas));
    if (action === 'deleteUploadData') return apiResponse(deleteUploadData(data.period, data.floor));
    if (action === 'uploadMultipleCsvsToDrive') return apiResponse(uploadMultipleCsvsToDrive(data.filesArray, data.period));
    if (action === 'clearDatabase') return apiResponse(clearDatabase());
    if (action === 'sendSummaryEmail') return apiResponse(sendSummaryEmail(data.period));

    return apiResponse({ success: false, message: "ไม่พบคำสั่ง (Action) ที่ระบุ: " + action });
  } catch (err) {
    return apiResponse({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  return ContentService.createTextOutput("✅ Print System Backend API is Running Properly! (V24.0)");
}

function apiResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. Authentication & User Management
// ==========================================
function loginUser(mobile, pin) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.users);
  if (!sh) return { success: false, message: "ไม่พบชีตผู้ใช้งาน" };
  var data = sh.getDataRange().getValues();
  
  var inputMobile = String(mobile).trim();
  var inputPin = String(pin).trim();
  
  for (var i = 1; i < data.length; i++) {
    var dbMobile = String(data[i][0]).replace(/'/g, '').trim();
    var dbPin = String(data[i][1]).replace(/'/g, '').trim();
    
    if (dbMobile === inputMobile && dbPin === inputPin) {
      if (data[i][7] === "pending") {
        return { success: false, message: "บัญชีของคุณรอการอนุมัติจาก Admin" };
      }
      if (data[i][7] === "suspended") {
        return { success: false, message: "บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อ Admin" };
      }
      return { 
        success: true, 
        user: { 
          mobile: dbMobile, 
          name: data[i][3], 
          role: data[i][6], 
          group: data[i][5], 
          printerUser: data[i][11] || "",
          email: data[i][13] || ""
        } 
      };
    }
  }
  return { success: false, message: "เบอร์มือถือ หรือ PIN ไม่ถูกต้อง" };
}

function registerUser(regData) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var inputMobile = String(regData.mobile).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
      return { success: false, message: "เบอร์มือถือนี้ลงทะเบียนไว้แล้ว" };
    }
  }
  
  var pin = Math.floor(100000 + Math.random() * 900000).toString();
  sh.appendRow([ "'" + inputMobile, pin, regData.gender || "", regData.name, regData.position || "", regData.group, "user", "pending", new Date(), 0, "", regData.printerUser || "", "", regData.email || "" ]);
  
  return { success: true, pin: pin, message: "สมัครสมาชิกสำเร็จ! กรุณาจดจำรหัส PIN ของท่าน" };
}

function adminCreateUser(regData) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var inputMobile = String(regData.mobile).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
      return { success: false, message: "เบอร์มือถือนี้มีในระบบแล้ว ไม่สามารถสร้างซ้ำได้" };
    }
  }
  
  var pin = Math.floor(100000 + Math.random() * 900000).toString();
  sh.appendRow([ "'" + inputMobile, pin, regData.gender || "", regData.name, regData.position || "", regData.group, regData.role || "user", "approved", new Date(), 0, "", regData.printerUser || "", "", regData.email || "" ]);
  
  return { success: true, pin: pin, message: "สร้างบัญชีผู้ใช้สำเร็จ! รหัส PIN ของผู้ใช้นี้คือ: " + pin };
}

function requestNewPin(mobile) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var currentMonth = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM");
  var inputMobile = String(mobile).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
      var pinCount = parseInt(data[i][9]) || 0;
      var pinMonth = String(data[i][10] || "");
      
      if (pinMonth === currentMonth) {
        if (pinCount >= 3) {
          return { success: false, message: "คุณขอเปลี่ยน PIN ครบ 3 ครั้งในเดือนนี้แล้ว กรุณาติดต่อ Admin" };
        }
        pinCount++;
      } else {
        pinMonth = currentMonth;
        pinCount = 1;
      }
      
      var newPin = Math.floor(100000 + Math.random() * 900000).toString();
      sh.getRange(i + 1, 2).setValue(newPin); 
      
      if (sh.getMaxColumns() < 12) {
        sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
      }
      
      sh.getRange(i + 1, 10).setValue(pinCount);
      sh.getRange(i + 1, 11).setValue(pinMonth);
      
      return { success: true, pin: newPin, message: "สร้างรหัส PIN ใหม่สำเร็จแล้ว: " + newPin };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลเบอร์มือถือนี้" };
}

function editUser(mobile, updatedData) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var inputMobile = String(mobile).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
      sh.getRange(i + 1, 1).setValue("'" + String(updatedData.mobile).trim());
      sh.getRange(i + 1, 3).setValue(updatedData.gender || "");
      sh.getRange(i + 1, 4).setValue(updatedData.name);
      sh.getRange(i + 1, 5).setValue(updatedData.position);
      sh.getRange(i + 1, 6).setValue(updatedData.group);
      sh.getRange(i + 1, 7).setValue(updatedData.role);
      sh.getRange(i + 1, 8).setValue(updatedData.status);
      sh.getRange(i + 1, 12).setValue(updatedData.printerUser || "");
      if (updatedData.email) sh.getRange(i + 1, 14).setValue(updatedData.email);
      return { success: true, message: "อัปเดตข้อมูลบัญชีผู้ใช้สำเร็จ" };
    }
  }
  return { success: false, message: "ไม่พบบัญชีผู้ใช้งานนี้ในระบบ" };
}

function suspendUser(mobile) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var inputMobile = String(mobile).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
      var currentStatus = String(data[i][7]);
      var newStatus = currentStatus === "suspended" ? "approved" : "suspended";
      var msg = newStatus === "suspended" ? "ระงับการใช้งานบัญชีสำเร็จ" : "ปลดระงับใช้งานสำเร็จ";
      
      sh.getRange(i + 1, 8).setValue(newStatus);
      return { success: true, message: msg };
    }
  }
  return { success: false, message: "ไม่พบบัญชีผู้ใช้งานนี้ในระบบ" };
}

function approveUser(mobile) {
  var sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var inputMobile = String(mobile).trim();
  for (var i = 1; i < data.length; i++) { 
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) { 
      sh.getRange(i + 1, 8).setValue("approved"); 
      return { success: true, message: "อนุมัติบัญชีสำเร็จ" }; 
    } 
  }
  return { success: false, message: "ไม่พบบัญชีผู้ใช้งานนี้" };
}

function deleteUser(mobile) {
  var sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEETS.users);
  var data = sh.getDataRange().getValues();
  var inputMobile = String(mobile).trim();
  for (var i = 1; i < data.length; i++) { 
    if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) { 
      sh.deleteRow(i + 1); 
      return { success: true, message: "ลบผู้ใช้สำเร็จ" }; 
    } 
  }
  return { success: false, message: "ไม่พบบัญชีผู้ใช้งานนี้" };
}

// ==========================================
// 3. Staff Verification & Self Profile Edit
// ==========================================
function getStaffByPrinterCode(printerCode) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var pCode = String(printerCode).trim();
    
    // Check in Staff_Map
    var shStaff = ss.getSheetByName(CONFIG.SHEETS.staff_map);
    var staffData = shStaff ? shStaff.getDataRange().getValues() : [];
    var foundStaff = null;
    
    for (var i = 1; i < staffData.length; i++) {
      var scode = String(staffData[i][0]).trim();
      if (scode === pCode) {
        foundStaff = {
          printerCode: scode,
          name: staffData[i][1],
          groupCode: String(staffData[i][2]).padStart(2, '0'),
          groupName: staffData[i][3]
        };
        break;
      }
    }
    
    if (!foundStaff) {
      return { success: false, message: "ไม่พบรหัสเครื่องปริ้นนี้ในระบบ กรุณาตรวจสอบอีกครั้ง" };
    }
    
    // Check if user has already registered in App_Users
    var shUsers = ss.getSheetByName(CONFIG.SHEETS.users);
    if (shUsers) {
      var uData = shUsers.getDataRange().getValues();
      for (var j = 1; j < uData.length; j++) {
        var pu = String(uData[j][11] || "").trim();
        if (pu === pCode) {
          foundStaff.mobile = String(uData[j][0]).replace(/'/g, '').trim();
          foundStaff.position = uData[j][4];
          foundStaff.email = uData[j][13] || "";
          break;
        }
      }
    }
    
    return {
      success: true,
      data: foundStaff,
      groups: getGroupOptions()
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function updateStaffData(payload) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var shUsers = ss.getSheetByName(CONFIG.SHEETS.users);
    var uData = shUsers.getDataRange().getValues();
    var pCode = String(payload.printerCode).trim();
    var mobile = String(payload.mobile).trim();
    var existingRow = -1;
    
    for (var i = 1; i < uData.length; i++) {
      var dbMobile = String(uData[i][0]).replace(/'/g, '').trim();
      var dbPCode = String(uData[i][11] || "").trim();
      if (dbMobile === mobile || (pCode && dbPCode === pCode)) {
        existingRow = i + 1;
        break;
      }
    }
    
    if (existingRow > 0) {
      // Update existing
      shUsers.getRange(existingRow, 1).setValue("'" + mobile);
      shUsers.getRange(existingRow, 4).setValue(payload.name);
      shUsers.getRange(existingRow, 5).setValue(payload.position);
      shUsers.getRange(existingRow, 6).setValue(payload.groupName);
      shUsers.getRange(existingRow, 12).setValue(pCode);
      if (shUsers.getMaxColumns() < 14) shUsers.insertColumnsAfter(shUsers.getMaxColumns(), 14 - shUsers.getMaxColumns());
      shUsers.getRange(existingRow, 14).setValue(payload.email || "");
      
      return { success: true, isNew: false, message: "อัปเดตข้อมูลของท่านเรียบร้อยแล้ว" };
    } else {
      // Create new user with PIN
      var pin = Math.floor(100000 + Math.random() * 900000).toString();
      shUsers.appendRow([
        "'" + mobile,
        pin,
        "",
        payload.name,
        payload.position,
        payload.groupName,
        "user",
        "approved",
        new Date(),
        0,
        "",
        pCode,
        "",
        payload.email || ""
      ]);
      
      return { success: true, isNew: true, pin: pin, message: "ยืนยันตัวตนและสร้างบัญชีสำเร็จ!" };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function requestProfileEdit(mobile, newData) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.users);
    var data = sh.getDataRange().getValues();
    var inputMobile = String(mobile).trim();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
        if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());
        sh.getRange(i + 1, 13).setValue(JSON.stringify(newData));
        
        // Notify Admin via Telegram
        var settingsMap = getSettingsMap_();
        var alertMsg = "📝 *มีคำขอแก้ไขข้อมูลส่วนตัวใหม่*\n" +
                       "👤 *ผู้ขอ:* " + data[i][3] + " (เบอร์เดิม: " + inputMobile + ")\n" +
                       "🔄 *ชื่อใหม่:* " + newData.name + "\n" +
                       "🏢 *กลุ่มงานใหม่:* " + newData.groupName + "\n" +
                       "📱 *เบอร์ใหม่:* " + newData.mobile + "\n" +
                       "โปรดตรวจสอบในแท็บจัดการแอดมิน";
        sendTelegramAlert_(alertMsg, settingsMap);
        
        return { success: true, message: "ส่งคำขอแก้ไขข้อมูลให้ Admin เรียบร้อยแล้ว" };
      }
    }
    return { success: false, message: "ไม่พบบัญชีผู้ใช้" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function approveProfileEdit(mobile, decision) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.users);
    var data = sh.getDataRange().getValues();
    var inputMobile = String(mobile).trim();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).replace(/'/g, '').trim() === inputMobile) {
        if (decision === 'approve') {
          var pendingJson = data[i][12];
          if (!pendingJson) return { success: false, message: "ไม่มีข้อมูลคำขอแก้ไขที่รออนุมัติ" };
          var newData = JSON.parse(pendingJson);
          
          sh.getRange(i + 1, 1).setValue("'" + String(newData.mobile).trim());
          sh.getRange(i + 1, 4).setValue(newData.name);
          sh.getRange(i + 1, 5).setValue(newData.position || "");
          sh.getRange(i + 1, 6).setValue(newData.groupName || data[i][5]);
          if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());
          sh.getRange(i + 1, 14).setValue(newData.email || "");
          sh.getRange(i + 1, 13).setValue(""); // Clear pending
          
          return { success: true, message: "อนุมัติคำขอแก้ไขข้อมูลสำเร็จ" };
        } else {
          sh.getRange(i + 1, 13).setValue(""); // Clear pending
          return { success: true, message: "ปฏิเสธคำขอแก้ไขข้อมูลเรียบร้อยแล้ว" };
        }
      }
    }
    return { success: false, message: "ไม่พบบัญชีผู้ใช้" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 4. Print Requests & Telegram Notifications
// ==========================================
function submitPrintRequest(reqData) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.print_req);
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_F1_ID);
    var reqFolder = getOrCreateSubfolder_(folder, "Color_Print_Requests");
    
    var fileUrls = [];
    var totalPages = 0;
    var filesArray = reqData.files || [];
    
    // Support single file or multi-file
    if (filesArray.length === 0 && reqData.fileDataURI) {
      filesArray.push({
        fileName: reqData.fileName,
        mimeType: reqData.mimeType,
        fileDataURI: reqData.fileDataURI,
        pages: reqData.pages,
        copies: reqData.copies || 1,
        pageRange: reqData.pageRange || "ทั้งหมด"
      });
    }
    
    var detailsList = [];
    filesArray.forEach(function(item) {
      var base64Data = item.fileDataURI.split(',')[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), item.mimeType, item.fileName);
      var file = reqFolder.createFile(blob);
      fileUrls.push(file.getUrl());
      var p = parseInt(item.pages) || 1;
      var c = parseInt(item.copies) || 1;
      totalPages += (p * c);
      detailsList.push(item.fileName + " (" + (item.pageRange || "ทั้งหมด") + ", " + p + " แผ่น x " + c + " ชุด)");
    });
    
    var timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    var period = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MMMM yyyy");
    var combinedFileUrl = fileUrls.join("\n");
    var detailsStr = detailsList.join("; ");
    
    sh.appendRow([
      timestamp,
      period,
      "'" + String(reqData.mobile).trim(),
      reqData.name,
      reqData.group,
      combinedFileUrl,
      totalPages,
      "รอปริ้น",
      detailsStr,
      reqData.note || ""
    ]);
    
    // Send Notifications (Email & Telegram)
    var settingsMap = getSettingsMap_();
    
    // 1. Email notification
    if (settingsMap.notify_emails) {
      var html = "<h3>มีคำขอปริ้นสีใหม่เข้าสู่ระบบ</h3>" +
                 "<p><strong>ผู้ขอ:</strong> " + reqData.name + " (" + reqData.group + ")</p>" +
                 "<p><strong>เบอร์โทรศัพท์:</strong> " + reqData.mobile + "</p>" +
                 "<p><strong>ยอดรวม:</strong> " + totalPages + " แผ่น</p>" +
                 "<p><strong>รายละเอียด:</strong> " + detailsStr + "</p>" +
                 "<p><strong>ไฟล์แนบ:</strong> " + fileUrls.map(function(u){ return "<a href='" + u + "'>เปิดไฟล์</a>"; }).join(", ") + "</p>";
      MailApp.sendEmail({ to: settingsMap.notify_emails, subject: "[คำขอปริ้นสีใหม่] จาก " + reqData.name, htmlBody: html });
    }
    
    // 2. Telegram notification
    var tgMsg = "🖨️ *มีคำขอปริ้นสีใหม่เข้ามา!*\n" +
                "👤 *ผู้ขอ:* " + reqData.name + "\n" +
                "🏢 *กลุ่มงาน:* " + reqData.group + "\n" +
                "📱 *เบอร์ติดต่อ:* " + reqData.mobile + "\n" +
                "📄 *จำนวน:* " + totalPages + " แผ่น\n" +
                "📋 *รายละเอียด:* " + detailsStr;
    sendTelegramAlert_(tgMsg, settingsMap);
    
    return { success: true, message: "ส่งคำขอและอัปโหลดไฟล์สำเร็จ ระบบได้แจ้งเตือน Admin เรียบร้อยแล้ว" };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

function updatePrintReqStatus(tsStr, newStatus) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.print_req);
    var data = sh.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var rTs = data[i][0];
      var rTsFmt = (rTs instanceof Date) ? Utilities.formatDate(rTs, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(rTs).trim();
      
      if (rTsFmt === String(tsStr).trim() || String(rTs).trim() === String(tsStr).trim()) {
        sh.getRange(i + 1, 8).setValue(newStatus);
        
        // Feature B: แจ้งเตือนเมื่อสถานะเปลี่ยนเป็น "เสร็จสิ้น" หรือ "ยกเลิก"
        var settingsMap = getSettingsMap_();
        var requesterName = data[i][3];
        var requesterGroup = data[i][4];
        var requesterMobile = String(data[i][2]).replace(/'/g, '').trim();
        var pages = data[i][6];
        
        if (newStatus === "เสร็จสิ้น") {
          var successMsg = "🎉 *คำขอปริ้นสีเสร็จสิ้นแล้ว!*\n" +
                           "👤 *ผู้ขอ:* " + requesterName + " (" + requesterGroup + ")\n" +
                           "📄 *จำนวน:* " + pages + " แผ่น\n" +
                           "🏢 *สถานที่รับ:* กลุ่มงานสุขภาพดิจิทัล (ชั้น 2 สสจ.นครนายก)\n" +
                           "✨ เจ้าหน้าที่สามารถมารับเอกสารได้เลยครับ";
          sendTelegramAlert_(successMsg, settingsMap);
        } else if (newStatus === "ยกเลิก") {
          var cancelMsg = "❌ *คำขอปริ้นสีถูกยกเลิก*\n" +
                          "👤 *ผู้ขอ:* " + requesterName + " (" + requesterGroup + ")\n" +
                          "📄 *จำนวน:* " + pages + " แผ่น\n" +
                          "กรุณาติดต่อสอบถาม Admin เพิ่มเติมครับ";
          sendTelegramAlert_(cancelMsg, settingsMap);
        }
        
        return { success: true, message: "อัปเดตสถานะเป็น '" + newStatus + "' สำเร็จ" };
      }
    }
    return { success: false, message: "ไม่พบคำขอนี้ในระบบ" };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

function deletePrintReq(tsStr) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.print_req);
    var data = sh.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var rTs = data[i][0];
      var rTsFmt = (rTs instanceof Date) ? Utilities.formatDate(rTs, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(rTs).trim();
      
      if (rTsFmt === String(tsStr).trim() || String(rTs).trim() === String(tsStr).trim()) {
        sh.deleteRow(i + 1);
        return { success: true, message: "ลบประวัติคำขอสำเร็จ" };
      }
    }
    return { success: false, message: "ไม่พบประวัติคำขอนี้ในระบบ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 5. Chat System
// ==========================================
function sendChatMessage(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = openOrCreateSheet_(CONFIG.SHEET_ID, CONFIG.SHEETS.chats);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Room_ID", "Chat_Time", "Sender_Name", "Sender_Role", "Message"]);
    }
    
    var timeStamp = new Date();
    sh.appendRow([
      "'" + String(data.roomId).trim(),
      timeStamp,
      data.senderName,
      data.senderRole,
      data.message
    ]);
    
    // If sent by user, notify Admin in Telegram
    if (data.senderRole === 'user') {
      var settingsMap = getSettingsMap_();
      var chatAlert = "💬 *ข้อความสนทนาใหม่จากผู้ใช้*\n" +
                      "👤 *ผู้ส่ง:* " + data.senderName + " (ห้อง: " + data.roomId + ")\n" +
                      "💬 *ข้อความ:* " + data.message;
      sendTelegramAlert_(chatAlert, settingsMap);
    }
    
    return { success: true, message: "ส่งข้อความสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function editChatMessage(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.chats);
    if (!sh) return { success: false, message: "ไม่พบชีตแชท" };
    var values = sh.getDataRange().getValues();
    
    for (var i = 1; i < values.length; i++) {
      var rid = String(values[i][0]).replace(/'/g, '').trim();
      var t = values[i][1];
      var tStr = (t instanceof Date) ? Utilities.formatDate(t, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(t).trim();
      
      if (rid === String(data.roomId).trim() && (tStr === data.chatTimeStr || String(t).trim() === data.chatTimeStr)) {
        if (values[i][2] === data.senderName) {
          sh.getRange(i + 1, 5).setValue(data.newMessage + " (แก้ไขแล้ว)");
          return { success: true, message: "แก้ไขข้อความสำเร็จ" };
        } else {
          return { success: false, message: "คุณไม่มีสิทธิ์แก้ไขข้อความของผู้อื่น" };
        }
      }
    }
    return { success: false, message: "ไม่พบข้อความที่คุณต้องการแก้ไข" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteChatMessage(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.chats);
    if (!sh) return { success: false, message: "ไม่พบชีตแชท" };
    var values = sh.getDataRange().getValues();
    
    for (var i = 1; i < values.length; i++) {
      var rid = String(values[i][0]).replace(/'/g, '').trim();
      var t = values[i][1];
      var tStr = (t instanceof Date) ? Utilities.formatDate(t, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(t).trim();
      
      if (rid === String(data.roomId).trim() && (tStr === data.chatTimeStr || String(t).trim() === data.chatTimeStr)) {
        if (values[i][2] === data.senderName || data.senderRole === 'admin') {
          sh.deleteRow(i + 1);
          return { success: true, message: "ลบข้อความสำเร็จ" };
        } else {
          return { success: false, message: "คุณไม่มีสิทธิ์ลบข้อความนี้" };
        }
      }
    }
    return { success: false, message: "ไม่พบข้อความที่คุณต้องการลบ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 6. Settings, Groups & Quota Configuration
// ==========================================
function saveSettings(costBw, costColor, emails, tgToken, tgChatId) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.settings);
  var data = sh.getDataRange().getValues();
  var foundKeys = {};
  
  for (var i = 1; i < data.length; i++) {
    var k = data[i][0];
    foundKeys[k] = i + 1;
    if (k === "cost_bw" && costBw !== undefined) sh.getRange(i + 1, 2).setValue(costBw);
    if (k === "cost_color" && costColor !== undefined) sh.getRange(i + 1, 2).setValue(costColor);
    if (k === "notify_emails" && emails !== undefined) sh.getRange(i + 1, 2).setValue(emails);
    if (k === "telegram_bot_token" && tgToken !== undefined) sh.getRange(i + 1, 2).setValue(tgToken);
    if (k === "telegram_chat_id" && tgChatId !== undefined) sh.getRange(i + 1, 2).setValue(tgChatId);
  }
  
  // Insert missing keys if any
  if (!foundKeys["telegram_bot_token"] && tgToken !== undefined) sh.appendRow(["telegram_bot_token", tgToken]);
  if (!foundKeys["telegram_chat_id"] && tgChatId !== undefined) sh.appendRow(["telegram_chat_id", tgChatId]);
  
  return { success: true, message: "บันทึกการตั้งค่าสำเร็จ" };
}

function saveGroupQuotas(quotas) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS.settings);
    var data = sh.getDataRange().getValues();
    var quotaJson = JSON.stringify(quotas || {});
    var updated = false;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "group_quotas") {
        sh.getRange(i + 1, 2).setValue(quotaJson);
        updated = true;
        break;
      }
    }
    if (!updated) {
      sh.appendRow(["group_quotas", quotaJson]);
    }
    return { success: true, message: "บันทึกโควตากลุ่มงานสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function saveGroupConfig(groups) {
  var sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEETS.groups);
  sh.clear();
  sh.appendRow(["รหัสกลุ่ม", "ชื่อกลุ่มงาน"]);
  if (groups.length > 0) {
    var mappedData = groups.map(function(g) { return ["'" + String(g.code).padStart(2, '0'), g.name]; });
    sh.getRange(2, 1, mappedData.length, 2).setValues(mappedData);
  }
  return { success: true };
}

function getAdminData() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  
  var users = ss.getSheetByName(CONFIG.SHEETS.users).getDataRange().getValues().slice(1).map(function(r) {
    return { 
      mobile: String(r[0]).replace(/'/g, '').trim(), 
      pin: String(r[1]).replace(/'/g, '').trim(), 
      gender: r[2],
      name: r[3], 
      position: r[4],
      group: r[5], 
      role: r[6], 
      status: r[7], 
      date: r[8], 
      printerUser: r[11],
      pendingEdit: r[12] || "",
      email: r[13] || ""
    };
  });
  
  var settingsMap = getSettingsMap_();
  var groups = getGroupOptions();
  
  return { users: users, settings: settingsMap, groups: groups };
}

function getDashboardData(userMobile) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  function getSheetData(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    var data = sh.getDataRange().getValues(); 
    var hdr = data.shift();
    return data.map(function(r) { 
      var obj = {}; 
      hdr.forEach(function(h, i) { obj[h] = r[i]; }); 
      return obj; 
    });
  }
  
  var settingsMap = getSettingsMap_();
  var printReqsRaw = getSheetData(CONFIG.SHEETS.print_req);
  printReqsRaw.forEach(function(r) {
    r["Timestamp_Str"] = (r["Timestamp"] instanceof Date) ? Utilities.formatDate(r["Timestamp"], CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(r["Timestamp"]);
  });

  var chatsRaw = getSheetData(CONFIG.SHEETS.chats);
  chatsRaw.forEach(function(c) {
    c["Chat_Time_Str"] = (c["Chat_Time"] instanceof Date) ? Utilities.formatDate(c["Chat_Time"], CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(c["Chat_Time"]);
  });

  // User profile
  var myProfile = null;
  if (userMobile) {
    var shUsers = ss.getSheetByName(CONFIG.SHEETS.users);
    if (shUsers) {
      var uData = shUsers.getDataRange().getValues();
      var inputM = String(userMobile).trim();
      for (var i = 1; i < uData.length; i++) {
        if (String(uData[i][0]).replace(/'/g, '').trim() === inputM) {
          myProfile = {
            mobile: inputM,
            name: uData[i][3],
            position: uData[i][4],
            groupName: uData[i][5],
            role: uData[i][6],
            pendingEdit: uData[i][12] || "",
            email: uData[i][13] || ""
          };
          break;
        }
      }
    }
  }

  return { 
    totals: getSheetData("agg_total"), 
    stacks: getSheetData("agg_all_stack"),
    staffs: getSheetData(CONFIG.SHEETS.agg_staff),
    settings: settingsMap,
    printReqs: printReqsRaw,
    chats: chatsRaw,
    groups: getGroupOptions(),
    myProfile: myProfile
  };
}

function getGroupOptions() {
  var groups = getGroupMap_();
  var options = [];
  for (var key in groups) {
    options.push({ code: String(key).padStart(2, '0'), name: groups[key] });
  }
  options.sort(function(a, b) { return a.code.localeCompare(b.code); });
  return options;
}

function getSettingsMap_() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.settings);
  var map = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getDataRange().getValues().slice(1).forEach(function(r) {
      map[r[0]] = r[1];
    });
  }
  return map;
}

function sendTelegramAlert_(text, settingsMap) {
  try {
    var token = (settingsMap && settingsMap.telegram_bot_token) || "";
    var chatId = (settingsMap && settingsMap.telegram_chat_id) || "";
    if (!token || !chatId) return;

    var url = "https://api.telegram.org/bot" + token + "/sendMessage";
    var payload = {
      chat_id: chatId,
      text: text,
      parse_mode: "Markdown"
    };
    
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log("Telegram Error: " + e.toString());
  }
}

// ==========================================
// 7. Core ETL & File Processing
// ==========================================
function uploadMultipleCsvsToDrive(filesArray, period) {
  try {
    filesArray.forEach(function(item) {
      var folderId = CONFIG["FOLDER_" + item.floorKey + "_ID"];
      var folder = DriveApp.getFolderById(folderId);
      var base64Data = item.dataURI.split(',')[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'text/csv', "[" + period + "] " + item.fileName);
      folder.createFile(blob);
    });
    runWeeklyETL(); 
    return { success: true, message: "อัปโหลดและประมวลผลข้อมูลทั้ง " + filesArray.length + " ชั้นสำเร็จ!" };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

function deleteUploadData(period, floor) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.SHEETS["raw_" + floor.toLowerCase()]);
    if (!sh) return { success: false, message: "ไม่พบฐานข้อมูลชั้นที่เลือก" };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "ไม่มีข้อมูลให้ลบ" };

    var newData = [];
    var deletedCount = 0;
    var targetPeriod = String(period).trim();
    newData.push(data[0]); 

    for (var i = 1; i < data.length; i++) {
      var rPeriod = data[i][0];
      var rPeriodStr = (rPeriod instanceof Date) ? Utilities.formatDate(rPeriod, CONFIG.TIMEZONE, "yyyy-MM-dd") : String(rPeriod).trim();
      if (rPeriodStr === targetPeriod || String(rPeriod).trim() === targetPeriod) {
        deletedCount++;
      } else {
        newData.push(data[i]);
      }
    }

    if (deletedCount > 0) {
      sh.clearContents();
      if (newData.length > 0) sh.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      aggregateAll_(); 
      return { success: true, message: "ลบข้อมูลสำเร็จจำนวน " + deletedCount + " รายการ และอัปเดตสถิติใหม่แล้ว" };
    } else {
      return { success: false, message: "ไม่พบข้อมูลรอบ [" + period + "] ในชั้นที่เลือก" };
    }
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

function clearDatabase() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheetsToClear = [
      CONFIG.SHEETS.raw_f1, CONFIG.SHEETS.raw_f2, CONFIG.SHEETS.raw_f3,
      CONFIG.SHEETS.agg_f1, CONFIG.SHEETS.agg_f2, CONFIG.SHEETS.agg_f3,
      "agg_all_stack", CONFIG.SHEETS.agg_all, "agg_total", CONFIG.SHEETS.agg_staff,
      CONFIG.SHEETS.print_req
    ];
    sheetsToClear.forEach(function(sheetName) {
      var sh = ss.getSheetByName(sheetName);
      if (sh && sh.getLastRow() > 1) {
        sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
      }
    });
    return { success: true, message: "รีเซ็ตฐานข้อมูลการปริ้นทั้งหมดเรียบร้อยแล้ว!" };
  } catch(e) {
    return { success: false, message: "เกิดข้อผิดพลาดในการล้างข้อมูล: " + e.toString() };
  }
}

function runWeeklyETL() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var staffSheet = ss.getSheetByName(CONFIG.SHEETS.staff_map);
  var staffDict = {};
  
  if (staffSheet && staffSheet.getLastRow() > 1) {
    var sData = staffSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      staffDict[String(sData[i][0]).trim()] = { name: sData[i][1], groupCode: String(sData[i][2]).padStart(2, '0'), groupName: sData[i][3] };
    }
  }

  var stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
  var groupMap = getGroupMap_();
  
  processFloor_("F1", CONFIG.FOLDER_F1_ID, CONFIG.SHEETS.raw_f1, stamp, staffDict, groupMap);
  processFloor_("F2", CONFIG.FOLDER_F2_ID, CONFIG.SHEETS.raw_f2, stamp, staffDict, groupMap);
  processFloor_("F3", CONFIG.FOLDER_F3_ID, CONFIG.SHEETS.raw_f3, stamp, staffDict, groupMap);

  aggregateAll_();
}

function processFloor_(floorLabel, folderId, rawSheetName, stamp, staffDict, groupMap) {
  var folder = DriveApp.getFolderById(folderId);
  var archive = getOrCreateSubfolder_(folder, "Archive");
  var files = folder.getFiles();
  var dataRows = [];

  while (files.hasNext()) {
    var file = files.next(); 
    var fname = file.getName();
    if (!/\.(csv)$/i.test(fname)) continue;
    
    var periodMatch = fname.match(/^\[(.*?)\]/);
    var period = periodMatch ? periodMatch[1] : "ไม่ระบุรอบข้อมูล";

    var parsed = parseCsvSmart_(file);
    if (!parsed || parsed.rows.length < 2) continue;

    var header = parsed.rows[0]; 
    var csv = parsed.rows;
    var nameIdx = detectNameColIndex_(header, csv.slice(1, Math.min(csv.length, 15)));
    var det = detectMetricColumns_(header);
    if (nameIdx < 0 || !det.printColor) continue;

    for (var i = 1; i < csv.length; i++) {
      var row = csv[i]; 
      if (!row || row.length === 0) continue;
      
      var user = safeGet_(row, nameIdx);
      CONFIG.RENAME_RULES.forEach(function(rule) { user = String(user).replace(rule.from, rule.to); });

      var pc = toNum_(safeGet_(row, det.printColor.index)); 
      var pbw = toNum_(safeGet_(row, det.printBW.index));
      var cc = toNum_(safeGet_(row, det.copyColor.index)); 
      var cbw = toNum_(safeGet_(row, det.copyBW.index));

      var code = ""; var staffName = "-"; var groupName = "";
      var matchId = String(user).match(/(\d{6})/); 
      var lookupKey = matchId ? matchId[1] : String(user).trim();

      if (staffDict[lookupKey]) { 
        code = staffDict[lookupKey].groupCode; staffName = staffDict[lookupKey].name; 
      } else if (staffDict[String(user).trim()]) { 
        code = staffDict[String(user).trim()].groupCode; staffName = staffDict[String(user).trim()].name; 
      } else { 
        code = extractGroupCode_(user); 
      }

      code = code ? String(code).padStart(2, '0') : "";
      groupName = groupMap[code] || code;

      dataRows.push(["'" + period, stamp, floorLabel, user, staffName, code, groupName, pc, pbw, cc, cbw]);
    }

    if (CONFIG.DO_ARCHIVE) {
      file.makeCopy(fname, archive); 
      folder.removeFile(file);
    }
  }

  if (dataRows.length > 0) {
    var sh = openOrCreateSheet_(CONFIG.SHEET_ID, rawSheetName);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["รอบข้อมูล (เดือน/ปี)","วันที่โหลด","ชั้น","Username ในเครื่อง","ชื่อ-สกุลบุคลากร","รหัสกลุ่ม","ชื่อกลุ่ม","Print สี","Print ขาวดำ","Copy สี","Copy ขาวดำ"]);
    }
    sh.getRange(sh.getLastRow() + 1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }
}

function aggregateAll_() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  
  function readSheet(name) {
    var sh = ss.getSheetByName(name); 
    if (!sh || sh.getLastRow() < 2) return {rows:[], idx:{}};
    var values = sh.getDataRange().getValues(); 
    var hdr = values[0]; 
    var idx = {}; 
    hdr.forEach(function(h, i) { idx[h] = i; });
    return { rows: values.slice(1), idx: idx };
  }
  
  var R1 = readSheet(CONFIG.SHEETS.raw_f1); var R2 = readSheet(CONFIG.SHEETS.raw_f2); var R3 = readSheet(CONFIG.SHEETS.raw_f3);
  function toNum(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }

  var outStack = openOrCreateSheet_(CONFIG.SHEET_ID, "agg_all_stack"); outStack.clear();
  outStack.appendRow(["รอบข้อมูล","ชั้น","รหัสกลุ่ม","ชื่อกลุ่ม","Print สี","Print ขาวดำ","รวม Print","Copy สี","Copy ขาวดำ","รวม Copy","รวมทั้งหมด (Print+Copy)"]);
  
  var stackMap = {};
  function addStack(data, floorLabel) {
    data.rows.forEach(function(r) {
      var i = data.idx; 
      var rawPeriod = r[i["รอบข้อมูล (เดือน/ปี)"]];
      var period = (rawPeriod instanceof Date) ? Utilities.formatDate(rawPeriod, CONFIG.TIMEZONE, "yyyy-MM-dd") : String(rawPeriod).trim();
      var code = String(r[i["รหัสกลุ่ม"]]).padStart(2, '0'); 
      var name = r[i["ชื่อกลุ่ม"]];
      var key = period + "|" + floorLabel + "|" + code;
      
      if (!stackMap[key]) stackMap[key] = { period: period, floor: floorLabel, code: code, name: name, pc: 0, pbw: 0, cc: 0, cbw: 0 };
      
      stackMap[key].pc += toNum(r[i["Print สี"]]); stackMap[key].pbw += toNum(r[i["Print ขาวดำ"]]);
      stackMap[key].cc += toNum(r[i["Copy สี"]]); stackMap[key].cbw += toNum(r[i["Copy ขาวดำ"]]);
    });
  }
  
  addStack(R1, "ชั้น 1"); addStack(R2, "ชั้น 2"); addStack(R3, "ชั้น 3");
  
  var stackRows = Object.keys(stackMap).map(function(k) {
    var o = stackMap[k]; 
    var sumP = o.pc + o.pbw; var sumC = o.cc + o.cbw;
    return ["'" + o.period, o.floor, "'" + o.code, o.name, o.pc, o.pbw, sumP, o.cc, o.cbw, sumC, sumP + sumC];
  });
  if (stackRows.length > 0) outStack.getRange(2, 1, stackRows.length, stackRows[0].length).setValues(stackRows);

  var outTotal = openOrCreateSheet_(CONFIG.SHEET_ID, "agg_total"); outTotal.clear();
  outTotal.appendRow(["รอบข้อมูล","รหัสกลุ่ม","ชื่อกลุ่ม","Print สี","Print ขาวดำ","รวม Print","Copy สี","Copy ขาวดำ","รวม Copy","รวมทั้งหมด (Print+Copy)"]);
  var totalMap = {};
  stackRows.forEach(function(r) {
    var period = String(r[0]).replace(/'/g, '').trim(); 
    var code = String(r[2]).replace(/'/g, '').padStart(2, '0'); 
    var name = r[3]; 
    var key = period + "|" + code;
    
    if (!totalMap[key]) totalMap[key] = { period: period, code: code, name: name, pc: 0, pbw: 0, cc: 0, cbw: 0 };
    totalMap[key].pc += r[4]; totalMap[key].pbw += r[5]; totalMap[key].cc += r[7]; totalMap[key].cbw += r[8];
  });
  
  var totalRows = Object.keys(totalMap).map(function(k) {
    var o = totalMap[k]; 
    var sumP = o.pc + o.pbw; var sumC = o.cc + o.cbw;
    return ["'" + o.period, "'" + o.code, o.name, o.pc, o.pbw, sumP, o.cc, o.cbw, sumC, sumP + sumC];
  });
  if (totalRows.length > 0) outTotal.getRange(2, 1, totalRows.length, totalRows[0].length).setValues(totalRows);

  var outStaff = openOrCreateSheet_(CONFIG.SHEET_ID, CONFIG.SHEETS.agg_staff); outStaff.clear();
  outStaff.appendRow(["รอบข้อมูล","รหัสกลุ่ม","ชื่อกลุ่ม","Username ในเครื่อง","ชื่อ-สกุลบุคลากร","Print สี","Print ขาวดำ","รวม Print","Copy สี","Copy ขาวดำ","รวม Copy","รวมทั้งหมด (Print+Copy)"]);
  var staffMap = {};
  
  function addStaff(data) {
    data.rows.forEach(function(r) {
      var i = data.idx; 
      var rawPeriod = r[i["รอบข้อมูล (เดือน/ปี)"]];
      var period = (rawPeriod instanceof Date) ? Utilities.formatDate(rawPeriod, CONFIG.TIMEZONE, "yyyy-MM-dd") : String(rawPeriod).trim();
      var user = r[i["Username ในเครื่อง"]]; 
      if (!user) return;
      
      var code = String(r[i["รหัสกลุ่ม"]]).padStart(2, '0'); 
      var gname = r[i["ชื่อกลุ่ม"]]; 
      var sname = r[i["ชื่อ-สกุลบุคลากร"]];
      var key = period + "|" + user;
      
      if (!staffMap[key]) staffMap[key] = { period: period, code: code, gname: gname, user: user, sname: sname, pc: 0, pbw: 0, cc: 0, cbw: 0 };
      staffMap[key].pc += toNum(r[i["Print สี"]]); staffMap[key].pbw += toNum(r[i["Print ขาวดำ"]]);
      staffMap[key].cc += toNum(r[i["Copy สี"]]); staffMap[key].cbw += toNum(r[i["Copy ขาวดำ"]]);
    });
  }
  
  addStaff(R1); addStaff(R2); addStaff(R3);
  var staffRows = Object.keys(staffMap).map(function(k) {
    var o = staffMap[k]; var sumP = o.pc + o.pbw; var sumC = o.cc + o.cbw;
    return ["'" + o.period, "'" + o.code, o.gname, "'" + o.user, o.sname, o.pc, o.pbw, sumP, o.cc, o.cbw, sumC, sumP + sumC];
  });
  staffRows.sort(function(a,b) { return b[11] - a[11]; }); 
  if (staffRows.length > 0) outStaff.getRange(2, 1, staffRows.length, staffRows[0].length).setValues(staffRows);
}

function sendSummaryEmail(period) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var shTotal = ss.getSheetByName("agg_total");
    var shSet = ss.getSheetByName(CONFIG.SHEETS.settings);
    
    var emailStr = ""; var costBw = 0.31; var costColor = 3;
    shSet.getDataRange().getValues().forEach(function(r) {
      if (r[0] === "notify_emails") emailStr = r[1];
      if (r[0] === "cost_bw") costBw = parseFloat(r[1]) || 0.31;
      if (r[0] === "cost_color") costColor = parseFloat(r[1]) || 3;
    });

    if (!emailStr) return { success: false, message: "ไม่ได้ตั้งค่าอีเมลปลายทางไว้ในระบบ" };

    var data = shTotal.getDataRange().getValues();
    var hdr = data.shift(); 
    var idx = {}; hdr.forEach(function(h, i) { idx[h] = i; });
    
    var filtered = data.filter(function(r) { return String(r[idx["รอบข้อมูล"]]).trim() === String(period).trim(); });
    if (filtered.length === 0) return { success: false, message: "ไม่มีข้อมูลในรอบที่เลือก" };

    var tUsage = 0, tCost = 0;
    filtered.sort(function(a, b) { return b[idx["รวมทั้งหมด (Print+Copy)"]] - a[idx["รวมทั้งหมด (Print+Copy)"]]; });
    filtered.forEach(function(r) { 
      tUsage += r[idx["รวมทั้งหมด (Print+Copy)"]]; 
      tCost += (r[idx["Print สี"]] + r[idx["Copy สี"]]) * costColor + (r[idx["Print ขาวดำ"]] + r[idx["Copy ขาวดำ"]]) * costBw; 
    });

    var html = "<div style='font-family: sans-serif; color: #333;'><h2 style='color: #2563eb;'>รายงานการใช้งานเครื่องถ่ายเอกสารและปริ้นเตอร์</h2><p><strong>รอบข้อมูล:</strong> " + period + "</p><div style='background: #f3f4f6; padding: 15px; border-radius: 8px; margin-bottom: 20px;'><h3 style='margin: 0 0 10px 0;'>ภาพรวม</h3><p>- ยอดใช้งานรวม: <strong>" + tUsage.toLocaleString() + "</strong> แผ่น</p><p>- ค่าใช้จ่ายประมาณการ: <strong style='color: #ef4444;'>฿" + tCost.toLocaleString() + "</strong></p></div><h3>5 อันดับกลุ่มงานที่ใช้งานสูงสุด</h3><table style='width: 100%; border-collapse: collapse; text-align: left;'><tr style='background: #e5e7eb;'><th style='padding: 8px; border: 1px solid #ddd;'>กลุ่มงาน</th><th style='padding: 8px; border: 1px solid #ddd;'>จำนวน (แผ่น)</th></tr>";
    
    filtered.slice(0, 5).forEach(function(r) { 
      html += "<tr><td style='padding: 8px; border: 1px solid #ddd;'>" + r[idx["ชื่อกลุ่ม"]] + "</td><td style='padding: 8px; border: 1px solid #ddd;'>" + r[idx["รวมทั้งหมด (Print+Copy)"]].toLocaleString() + "</td></tr>"; 
    });
    
    html += "</table><br><p>ท่านสามารถดูรายละเอียดเพิ่มเติมได้ที่ระบบจัดการผ่าน Web App</p></div>";
    MailApp.sendEmail({ to: emailStr, subject: "[สสจ.นครนายก] รายงานสถิติ Print & Copy รอบ " + period, htmlBody: html });
    return { success: true, message: "ส่งอีเมลสำเร็จ" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// Helpers
function getGroupMap_() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.SHEETS.groups);
  var map = {};
  if (sh && sh.getLastRow() > 1) {
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var code = String(data[i][0]).replace(/'/g, '').trim().padStart(2, '0');
      var name = data[i][1];
      map[code] = name;
    }
  }
  return map;
}

function parseCsvSmart_(file) {
  var txt = file.getBlob().getDataAsString(); 
  var rows = Utilities.parseCsv(txt);
  if (rows && rows[0] && rows[0].length > 1) return { rows: rows, delim: "," };
  
  var txt2 = txt.replace(/\r\n/g, "\n"); 
  var lines = txt2.split("\n");
  var rows2 = lines.map(function(line) { return line.split(";"); }); 
  if (rows2 && rows2[0] && rows2[0].length > 1) return { rows: rows2, delim: ";" };
  
  var rows3 = lines.map(function(line) { return line.split("\t"); }); 
  if (rows3 && rows3[0] && rows3[0].length > 1) return { rows: rows3, delim: "\\t" };
  return { rows: rows, delim: "," };
}

function detectNameColIndex_(h, sRows) {
  var c = ["User", "User Name", "Username", "Account", "Shared Account", "ผู้ใช้", "บัญชีผู้ใช้"];
  for (var i = 0; i < h.length; i++) { if (c.indexOf(String(h[i]).trim()) !== -1) return i; }
  
  var best = { idx: -1, score: -1 };
  for (var x = 0; x < h.length; x++) {
    var hits = 0;
    for (var r = 0; r < sRows.length; r++) { 
      var v = String(sRows[r][x] || ""); 
      if (/^\d{2}/.test(v)) hits += 2; 
      if (/_/.test(v)) hits += 1; 
    }
    if (hits > best.score) best = { idx: x, score: hits };
  }
  return best.idx;
}

function normalize_(s) { return String(s).toLowerCase().replace(/\s+/g, ""); }

function detectMetricColumns_(h) {
  var norm = h.map(function(x) { return normalize_(x); });
  function fi(pats) { 
    for (var i = 0; i < norm.length; i++) { 
      for (var p = 0; p < pats.length; p++) { if (pats[p].test(norm[i])) return { index: i }; } 
    } 
    return null; 
  }
  return {
    printColor: fi([/printer.*(fullcolour|colour|color|col|สี)/, /(สี).*printer/, /print.*(color|colour|สี)/, /printc/]),
    printBW: fi([/printer.*(black&white|bw|black|mono|ขาวดำ|ดำ)/, /(ดำ|ขาวดำ).*printer/, /printbw/]),
    copyColor: fi([/copy.*(fullcolour|colour|color|col|สี)/, /(สี).*copy/, /copyc/]),
    copyBW: fi([/copy.*(black&white|bw|black|mono|ขาวดำ|ดำ)/, /(ดำ|ขาวดำ).*copy/, /copybw/])
  };
}

function openOrCreateSheet_(id, name) { 
  var ss = SpreadsheetApp.openById(id); 
  var sh = ss.getSheetByName(name); 
  if (!sh) sh = ss.insertSheet(name); 
  return sh; 
}

function getOrCreateSubfolder_(f, name) { 
  var it = f.getFoldersByName(name); 
  if (it.hasNext()) return it.next();
  else return f.createFolder(name); 
}

function safeGet_(row, idx) { return (idx >= 0 && idx < row.length) ? String(row[idx]) : ""; }
function toNum_(v) { var s = String(v || "").replace(/,/g, "").trim(); var n = parseFloat(s); return isNaN(n) ? 0 : n; }
function extractGroupCode_(u) { var m = String(u).match(/^(\d{2})/); return m ? m[1] : ""; }

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === "runWeeklyETL") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("runWeeklyETL").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(18).create();
}

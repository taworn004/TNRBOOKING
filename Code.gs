function doGet() { 
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TNR Meeting Room Booking System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password); 
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) { 
    let hashVal = rawHash[i]; 
    if (hashVal < 0) hashVal += 256; 
    if (hashVal.toString(16).length == 1) txtHash += '0'; 
    txtHash += hashVal.toString(16); 
  } 
  return txtHash;
}

function loginUser(formObject) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues();
  const hashedPw = hashPassword(formObject.loginPassword);
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === formObject.loginUsername) {
      if (data[i][3] === hashedPw) return { success: true, message: "✅ เข้าสู่ระบบสำเร็จ!", userData: { uid: data[i][0], name: data[i][1], username: data[i][2], role: data[i][4] } };
      else return { success: false, message: "❌ รหัสผ่านไม่ถูกต้อง" };
    }
  } return { success: false, message: "❌ ไม่พบ Username นี้" };
}

function registerUser(formObject) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); 
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][2] === formObject.regUsername) return { success: false, message: "❌ Username นี้มีคนใช้งานแล้ว" }; }
  sheet.appendRow(["USR" + new Date().getTime(), formObject.regName, formObject.regUsername, hashPassword(formObject.regPassword), "User", ""]);
  return { success: true, message: "✅ ลงทะเบียนสำเร็จ! กรุณาเข้าสู่ระบบ" };
}

function uploadFileToDrive(base64Data, fileName) {
  if (!base64Data) return ""; 
  try {
    const folders = DriveApp.getFoldersByName("Meeting_Attachments"); 
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Meeting_Attachments");
    const split = base64Data.split(',');
    const contentType = split[0].split(';')[0].replace('data:', '');
    const decodedData = Utilities.base64Decode(split[1]);
    const blob = Utilities.newBlob(decodedData, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
    return file.getUrl();
  } catch (e) { return ""; }
}

function syncToCalendarAndEmail(bookingData, bookingId, fileUrl) {
  const { name, room, date, startTime, endTime, attendees, title, description } = bookingData;
  let desc = `📌 รหัส: ${bookingId}\n📝 หัวข้อ: ${title}\n👤 ผู้จอง: ${name}\n🏢 ห้อง: ${room}\nℹ️ รายละเอียด: \n${description}`;
  if (fileUrl) desc += `\n📄 ไฟล์แนบ: ${fileUrl}`;
  
  let eventId = "";
  try { eventId = CalendarApp.getDefaultCalendar().createEvent(`${title} (${room})`, new Date(`${date}T${startTime}:00`), new Date(`${date}T${endTime}:00`), { description: desc, guests: attendees, sendInvites: true }).getId(); } catch (e) {}

  if (attendees && attendees.trim() !== "") {
    try {
      let attachments = [];
      if (fileUrl) {
        let fileMatch = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (fileMatch) attachments.push(DriveApp.getFileById(fileMatch[1]).getBlob());
      }
      let htmlBody = `<div style="font-family:sans-serif; max-width:600px; border:1px solid #eee; border-radius:15px; overflow:hidden;"><div style="background:#4facfe; padding:20px; color:white; text-align:center;"><h2>✅ อนุมัติการจองห้องแล้ว</h2></div><div style="padding:20px;"><p><b>📌 หัวข้อ:</b> ${title}</p><p><b>📅 วันที่:</b> ${date}</p><p><b>⏰ เวลา:</b> ${startTime}-${endTime}</p><p><b>🏢 ห้อง:</b> ${room}</p><p><b>👤 ผู้จอง:</b> ${name}</p><p><b>ℹ️ รายละเอียด:</b><br>${description.replace(/\n/g, '<br>')}</p>${fileUrl ? `<p style="text-align:center; margin-top:20px;"><a href="${fileUrl}" style="background:#ff758c; color:white; padding:10px 20px; text-decoration:none; border-radius:10px;">📄 ดูไฟล์แนบ</a></p>` : ''}</div></div>`;
      GmailApp.sendEmail(attendees, `[ยืนยันการจองห้อง] ${title}`, "", { htmlBody: htmlBody, attachments: attachments, name: "TNR IT System" });
    } catch (e) {}
  }
  return eventId;
}

function getHREmails() {
  const users = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues();
  let hrEmails = [];
  for (let i = 1; i < users.length; i++) {
    if (users[i][4] === "HRManager" && users[i][2].includes("@")) hrEmails.push(users[i][2]);
  }
  return hrEmails;
}

function submitBooking(bookingData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings"); 
  const data = sheet.getDataRange().getValues();
  const { username, name, room, startDate, endDate, startTime, endTime, attendees, fileBase64, fileName, title, description, managerEmail } = bookingData;
  
  if (startTime >= endTime) return { success: false, message: "❌ เวลาเริ่มต้นต้องน้อยกว่าเวลาสิ้นสุด" };
  let start = new Date(startDate); let end = new Date(endDate);
  if (end < start) return { success: false, message: "❌ วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น" };
  
  let datesToBook = []; let current = new Date(start);
  while (current <= end) { datesToBook.push(Utilities.formatDate(new Date(current), Session.getScriptTimeZone(), "yyyy-MM-dd")); current.setDate(current.getDate() + 1); }

  for (let checkDate of datesToBook) {
    for (let i = 1; i < data.length; i++) {
      if (["Confirmed", "Pending_Dept", "Pending_HR"].includes(data[i][9])) {
        let bDate = (data[i][3] instanceof Date) ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][3]);
        if (data[i][2] === room && bDate === checkDate) {
          let bStart = (data[i][4] instanceof Date) ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), "HH:mm") : String(data[i][4]).substring(0, 5);
          let bEnd = (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "HH:mm") : String(data[i][5]).substring(0, 5);
          if (startTime < bEnd && endTime > bStart) return { success: false, message: `❌ คิวชนในวันที่ ${checkDate} (${bStart}-${bEnd})` };
        }
      }
    }
  }
  
  let fileUrl = uploadFileToDrive(fileBase64, fileName);
  let baseId = new Date().getTime();
  for (let i = 0; i < datesToBook.length; i++) {
    sheet.appendRow(["BK" + (baseId + i), new Date(), room, datesToBook[i], startTime, endTime, name, attendees, fileUrl, "Pending_Dept", "", title, description, managerEmail]);
  }

  if(managerEmail) {
    let subject = `[รออนุมัติขั้นที่ 1] คำขอใช้ห้องประชุมจาก ${name}`;
    let body = `เรียน หัวหน้าแผนก,\n\nมีคำขอจองห้องประชุมใหม่ รอการอนุมัติจากคุณ\nผู้จอง: ${name}\nหัวข้อ: ${title}\nห้อง: ${room}\nวันที่: ${datesToBook.join(', ')}\nเวลา: ${startTime}-${endTime}\n\nกรุณาเข้าสู่ระบบ TNR IT Dashboard เพื่อตรวจสอบและอนุมัติครับ`;
    try { GmailApp.sendEmail(managerEmail, subject, body, {name: "TNR System"}); } catch(e) {}
  }

  return { success: true, message: `✅ สำเร็จ! ส่งอีเมลแจ้งขออนุมัติไปยังหัวหน้าแผนกแล้ว` };
}

// 🟢 1. แก้ไขให้ HR มองเห็นรายการ Pending_Dept ด้วย แต่เอาไว้แสดงเฉยๆ
function getPendingApprovals(role, reqUsername) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings").getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(row => {
    if (role === "Admin") return ["Pending_Dept", "Pending_HR"].includes(row[9]);
    // ให้ HRManager เห็นทั้งที่รอแผนกและรอ HR
    if (role === "HRManager") return ["Pending_Dept", "Pending_HR"].includes(row[9]); 
    if (role === "DeptManager") return row[9] === "Pending_Dept" && row[13] === reqUsername; 
    return false;
  });
}

// 🟢 2. เพิ่มการป้องกันขั้นเด็ดขาด ไม่ให้ HR แอบกดยิงผ่าน API ได้ ถ้าระบบยังรอแผนกอยู่
function processApproval(bookingId, action, approverRole, approverUsername) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings"); const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookingId) {
      let savedManagerEmail = data[i][13]; 
      let currentStatus = data[i][9];

      // 🔴 ป้องกัน: ถ้ายังรอแผนกอยู่ HR ไม่มีสิทธิ์กดอนุมัติ/ปฏิเสธเด็ดขาด
      if (currentStatus === "Pending_Dept" && approverRole === "HRManager") {
        return { success: false, message: "❌ ต้องรอให้หัวหน้าแผนกอนุมัติก่อนครับ" };
      }
      if (currentStatus === "Pending_Dept" && approverRole === "DeptManager" && approverUsername !== savedManagerEmail) {
        return { success: false, message: "❌ คุณไม่มีสิทธิ์ (ไม่ใช่ผู้จัดการแผนกของคิวนี้)" };
      }

      if (action === "Reject") { sheet.getRange(i + 1, 10).setValue("Rejected"); return { success: true, message: "❌ ปฏิเสธรายการเรียบร้อย" }; }
      
      if (action === "Approve") {
        if (currentStatus === "Pending_Dept" && (approverRole === "DeptManager" || approverRole === "Admin")) {
          sheet.getRange(i + 1, 10).setValue("Pending_HR");
          let hrEmails = getHREmails();
          if (hrEmails.length > 0) {
            let dateStr = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
            let subject = `[รออนุมัติขั้นสุดท้าย] คำขอใช้ห้องประชุม: ${data[i][11]}`;
            let body = `เรียน ฝ่ายบุคคล (HR),\n\nคำขอจองห้องประชุมได้รับการอนุมัติจากหัวหน้าแผนกแล้ว รอพิจารณาขั้นสุดท้ายครับ\nผู้จอง: ${data[i][6]}\nห้อง: ${data[i][2]}\nวันที่: ${dateStr}\n\nกรุณาเข้าสู่ระบบเพื่อดำเนินการครับ`;
            try { GmailApp.sendEmail(hrEmails.join(","), subject, body, {name: "TNR System"}); } catch(e) {}
          }
          return { success: true, message: "✅ อนุมัติขั้นที่ 1 แล้ว! ระบบส่งอีเมลแจ้งฝ่ายบุคคลเรียบร้อย" };

        } else if (currentStatus === "Pending_HR" || (currentStatus === "Pending_Dept" && approverRole === "Admin")) {
          sheet.getRange(i + 1, 10).setValue("Confirmed");
          let bookingData = { name: data[i][6], room: data[i][2], date: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd"), startTime: Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), "HH:mm"), endTime: Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "HH:mm"), attendees: data[i][7], title: data[i][11], description: data[i][12] };
          sheet.getRange(i + 1, 11).setValue(syncToCalendarAndEmail(bookingData, bookingId, data[i][8]));
          return { success: true, message: "🎉 อนุมัติขั้นสุดท้ายสำเร็จ! ผู้จองได้รับอีเมลยืนยันแล้ว" };
        }
      }
    }
  } return { success: false, message: "❌ ขัดข้อง ไม่พบข้อมูล" };
}

function getBookingsList() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings").getDataRange().getDisplayValues().slice(1); }

function getCalendarEvents() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings").getDataRange().getValues(); let events = [];
  const colors = { "Room A (4 ที่นั่ง)": "#0d6efd", "Room B (10 ที่นั่ง)": "#198754", "Room C (20 ที่นั่ง)": "#dc3545" }; 
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === "Confirmed") {
      let bDate = (data[i][3] instanceof Date) ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][3]);
      let bStart = (data[i][4] instanceof Date) ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), "HH:mm") : String(data[i][4]).substring(0, 5);
      let bEnd = (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "HH:mm") : String(data[i][5]).substring(0, 5);
      events.push({ title: data[i][11] || data[i][2], start: `${bDate}T${bStart}:00`, end: `${bDate}T${bEnd}:00`, color: colors[data[i][2]] || "#6c757d", extendedProps: { booker: data[i][6], room: data[i][2], desc: data[i][12] || '-' } });
    }
  } return events;
}

function cancelBooking(bookingId, reqName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings"); const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookingId && data[i][6] === reqName) {
      sheet.getRange(i + 1, 10).setValue("Cancelled");
      if (data[i][10]) try { CalendarApp.getDefaultCalendar().getEventById(data[i][10]).deleteEvent(); } catch(e) {}
      return { success: true, message: "🗑️ ยกเลิกเรียบร้อย" };
    }
  } return { success: false, message: "❌ ไม่มีสิทธิ์" };
}

function adminGetUsersList(role) { return role === "Admin" ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getDisplayValues().slice(1) : []; }

function adminSaveUser(userData, reqRole) {
  if (reqRole !== "Admin") return { success: false };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); 
  const data = sheet.getDataRange().getValues();
  
  if (userData.uid === "") { 
    sheet.appendRow(["USR" + new Date().getTime(), userData.name, userData.username, hashPassword(userData.password), userData.role, userData.signature || ""]); 
    return { success: true, message: "✅ เพิ่มผู้ใช้และลายเซ็นแล้ว" }; 
  } else { 
    for (let i = 1; i < data.length; i++) { 
      if (data[i][0] === userData.uid) { 
        sheet.getRange(i+1, 2).setValue(userData.name); 
        sheet.getRange(i+1, 3).setValue(userData.username); 
        sheet.getRange(i+1, 5).setValue(userData.role); 
        sheet.getRange(i+1, 6).setValue(userData.signature || ""); 
        if (userData.password) sheet.getRange(i+1, 4).setValue(hashPassword(userData.password)); 
        return { success: true, message: "✅ อัปเดตข้อมูลผู้ใช้และลายเซ็นแล้ว" }; 
      } 
    } 
  }
}

function getSignatures() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let sigMap = {};
  
  for (let i = 1; i < data.length; i++) {
    let name = data[i][1];
    let email = data[i][2];
    let role = data[i][4];
    let sigUrl = data[i][5] || "https://cdn-icons-png.flaticon.com/512/3771/3771278.png"; 
    
    sigMap[name] = sigUrl; 
    sigMap[email] = sigUrl; 
    if (role === "HRManager") sigMap["HR_ADMIN"] = sigUrl; 
  }
  return sigMap;
}

function adminDeleteUser(uid, role) {
  if (role !== "Admin") return { success: false };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === uid && data[i][4] !== "Admin") { sheet.deleteRow(i + 1); return { success: true, message: "🗑️ ลบแล้ว" }; } }
  return { success: false, message: "❌ ลบไม่ได้" };
}

function getRoomsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName("Rooms");
  if (!sheet) { sheet = ss.insertSheet("Rooms"); sheet.appendRow(["ID", "Room Name", "Description"]); sheet.appendRow(["RM1", "Room A (4 ที่นั่ง)", "ห้องเล็ก"]); }
  return sheet.getDataRange().getValues().slice(1);
}

function saveRoom(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName("Rooms"); const values = sheet.getDataRange().getValues();
  if (data.id) { for (let i = 1; i < values.length; i++) { if (values[i][0] == data.id) { sheet.getRange(i + 1, 2, 1, 2).setValues([[data.name, data.desc]]); break; } } } 
  else { sheet.appendRow(["RM" + new Date().getTime(), data.name, data.desc]); } return { success: true };
}

function deleteRoom(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName("Rooms"); const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) { if (values[i][0] == id) { sheet.deleteRow(i + 1); break; } } return { success: true };
}

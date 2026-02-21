// ============================================
// ระบบคิวรับ Case กลาง - Google Apps Script
// ============================================

// ชื่อ Sheet
const DOCTORS_SHEET = "Doctors";
const CASES_SHEET = "Cases";

// ============================================
// 1) ฟังก์ชันตั้งค่าเริ่มต้น (รันครั้งเดียว)
// ============================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // สร้าง Sheet "Doctors" ถ้ายังไม่มี
  let doctorSheet = ss.getSheetByName(DOCTORS_SHEET);
  if (!doctorSheet) {
    doctorSheet = ss.insertSheet(DOCTORS_SHEET);
    doctorSheet.getRange("A1:C1").setValues([["ลำดับ", "ชื่อแพทย์", "สถานะ"]]);
    doctorSheet.getRange("A2:C6").setValues([
      [1, "พญ.สมศรี", "Active"],
      [2, "นพ.วิชัย", "Active"],
      [3, "พญ.นภา", "Active"],
      [4, "นพ.ธนา", "Active"],
      [5, "พญ.มณี", "Active"]
    ]);
    doctorSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    doctorSheet.setColumnWidth(2, 200);
  }
  
  // สร้าง Sheet "Cases" ถ้ายังไม่มี
  let caseSheet = ss.getSheetByName(CASES_SHEET);
  if (!caseSheet) {
    caseSheet = ss.insertSheet(CASES_SHEET);
    caseSheet.getRange("A1:G1").setValues([["Timestamp", "HN", "Diagnosis", "Ward", "แพทย์รับ Case", "ลำดับคิว", "หมายเหตุ"]]);
    caseSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    caseSheet.setColumnWidth(1, 180);
    caseSheet.setColumnWidth(2, 120);
    caseSheet.setColumnWidth(3, 200);
    caseSheet.setColumnWidth(4, 120);
    caseSheet.setColumnWidth(5, 160);
    caseSheet.setColumnWidth(6, 100);
  }
  
  // สร้าง Trigger สำหรับ onEdit (ลบตัวเก่าก่อน)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "onSheetEdit") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("onSheetEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  SpreadsheetApp.getUi().alert("✅ ตั้งค่าเรียบร้อย! สามารถเริ่มกรอก Case ได้เลยค่ะ");
}

// ============================================
// 2) ดึงรายชื่อแพทย์ที่ Active
// ============================================
function getActiveDoctors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DOCTORS_SHEET);
  const data = sheet.getDataRange().getValues();
  
  const doctors = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === "Active") {
      doctors.push({
        order: data[i][0],
        name: data[i][1],
        status: data[i][2]
      });
    }
  }
  // เรียงตามลำดับ
  doctors.sort((a, b) => a.order - b.order);
  return doctors;
}

// ============================================
// 3) นับจำนวน Case ของแต่ละแพทย์
// ============================================
function getCaseCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const doctor = data[i][4]; // คอลัมน์ "แพทย์รับ Case"
    if (doctor) {
      counts[doctor] = (counts[doctor] || 0) + 1;
    }
  }
  return counts;
}

// ============================================
// 4) หาแพทย์คนถัดไปตามคิว (Round-Robin)
// ============================================
function getNextDoctor() {
  const doctors = getActiveDoctors();
  if (doctors.length === 0) return null;
  
  const counts = getCaseCounts();
  
  // หาจำนวน case น้อยที่สุด
  let minCount = Infinity;
  doctors.forEach(doc => {
    const count = counts[doc.name] || 0;
    if (count < minCount) minCount = count;
  });
  
  // หาแพทย์ที่มี case น้อยที่สุด (ถ้าเท่ากัน เอาคนที่ลำดับน้อยกว่า)
  for (const doc of doctors) {
    const count = counts[doc.name] || 0;
    if (count === minCount) {
      return doc.name;
    }
  }
  
  return doctors[0].name;
}

// ============================================
// 5) Trigger เมื่อกรอกข้อมูลใน Cases Sheet
// ============================================
function onSheetEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CASES_SHEET) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // ถ้ากรอกคอลัมน์ B (HN) ในแถวใหม่
  if (col === 2 && row > 1) {
    const hn = sheet.getRange(row, 2).getValue();
    const existingDoctor = sheet.getRange(row, 5).getValue();
    
    // ถ้ามี HN แต่ยังไม่มีแพทย์
    if (hn && !existingDoctor) {
      const nextDoctor = getNextDoctor();
      const totalCases = getTotalCaseCount();
      
      // กรอก Timestamp
      if (!sheet.getRange(row, 1).getValue()) {
        sheet.getRange(row, 1).setValue(new Date());
        sheet.getRange(row, 1).setNumberFormat("dd/MM/yyyy HH:mm");
      }
      
      // กรอกแพทย์และลำดับคิว
      sheet.getRange(row, 5).setValue(nextDoctor);
      sheet.getRange(row, 6).setValue(totalCases + 1);
      
      // ไฮไลท์แถว
      highlightRow(sheet, row, nextDoctor);
    }
  }
}

// ============================================
// 6) ไฮไลท์สีตามแพทย์
// ============================================
function highlightRow(sheet, row, doctorName) {
  const doctors = getActiveDoctors();
  const colors = ["#E8F5E9", "#E3F2FD", "#FFF3E0", "#F3E5F5", "#FFEBEE", "#E0F7FA", "#FFF8E1"];
  
  const idx = doctors.findIndex(d => d.name === doctorName);
  const color = colors[idx % colors.length];
  
  sheet.getRange(row, 1, 1, 7).setBackground(color);
}

// ============================================
// 7) นับ Case ทั้งหมด
// ============================================
function getTotalCaseCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) count++; // นับแถวที่มี HN
  }
  return count;
}

// ============================================
// PDPA: ฟังก์ชันบดบัง HN (Masking)
// ============================================
function maskHN(hn) {
  if (!hn) return "";
  const str = String(hn);
  if (str.length <= 3) return "***";
  // แสดงแค่ 3 ตัวท้าย เช่น 6801234 → ****234
  const visible = str.slice(-3);
  const masked = "*".repeat(str.length - 3);
  return masked + visible;
}

// บดบัง HN ใน case array
function maskCases(cases) {
  return cases.map(c => ({
    ...c,
    hn: maskHN(c.hn)
  }));
}

// ============================================
// 8) Web App — Router
// ============================================
// URL Patterns:
//   ?page=form           → หน้ากรอก Case (พยาบาล)
//   ?page=api&payload={} → Form API (login, submit)
//   ?action=status       → JSON API (สำหรับ Dashboard website)
// ============================================
function doGet(e) {
  const page = e.parameter.page || "";
  
  // ---- Nurse Form Page: ฟอร์มย้ายไปอยู่คู่กับ queue-website แล้ว (ใช้ลิงก์จากหน้าแสดงคิว) ----
  if (page === "form") {
    const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>กรอก Case กลาง</title></head><body style="font-family:Sarabun,sans-serif;padding:24px;text-align:center;background:#0f172a;color:#f1f5f9;"><p style="font-size:18px;">ฟอร์มพยาบาลย้ายไปอยู่คู่กับหน้าแสดงคิวแล้ว</p><p style="margin-top:12px;">กรุณาเปิด <strong>หน้าแสดงคิว</strong> (ระบบคิวรับ Case กลาง) แล้วกดลิงก์<br>「📝 กรอก Case กลาง (ฟอร์มพยาบาล)」</p></body></html>';
    return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
  }

  // ---- Form API (called from NurseForm) ----
  if (page === "api") {
    try {
      const payload = JSON.parse(e.parameter.payload || "{}");
      const result = handleFormRequest(JSON.stringify(payload));
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // ---- JSON API (for Dashboard website) ----
  const action = e.parameter.action || "status";
  const mode = e.parameter.mode || "public";
  const token = e.parameter.token || "";
  
  const isInternal = (mode === "internal" && token === getInternalToken());
  
  let result;
  
  switch (action) {
    case "status":
      result = getQueueStatus();
      if (!isInternal) {
        result.recentCases = maskCases(result.recentCases);
      }
      break;
    case "doctors":
      result = getActiveDoctors();
      break;
    case "cases":
      result = getRecentCases(parseInt(e.parameter.limit) || 20);
      if (!isInternal) {
        result = maskCases(result);
      }
      break;
    case "search":
      if (!isInternal) {
        result = { error: "🔒 ไม่อนุญาตให้ค้นหา HN ในโหมดสาธารณะ (PDPA)" };
      } else {
        result = searchByHN(e.parameter.hn || "");
      }
      break;
    default:
      result = { error: "Unknown action" };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ดึง Internal Token จาก Script Properties
function getInternalToken() {
  try {
    return PropertiesService.getScriptProperties().getProperty("INTERNAL_TOKEN") || "changeme";
  } catch (e) {
    return "changeme";
  }
}

// ดึง Form Password จาก Script Properties
function getFormPassword() {
  try {
    return PropertiesService.getScriptProperties().getProperty("FORM_PASSWORD") || "nurse1234";
  } catch (e) {
    return "nurse1234";
  }
}

// ============================================
// Form API Handler (เรียกจาก NurseForm.html)
// ============================================
function handleFormRequest(payloadJson) {
  const payload = JSON.parse(payloadJson);
  const action = payload.action;
  
  switch (action) {
    case "login":
      return handleLogin(payload);
    case "status":
      return handleFormStatus(payload);
    case "submit":
      return handleSubmitCase(payload);
    default:
      return { error: "Unknown action" };
  }
}

// ---- Login ----
function handleLogin(payload) {
  const password = payload.password || "";
  const correctPw = getFormPassword();
  
  if (password === correctPw) {
    // สร้าง simple session token
    const token = Utilities.getUuid();
    // เก็บ token ใน Cache (หมดอายุ 8 ชั่วโมง)
    CacheService.getScriptCache().put("session_" + token, "valid", 28800);
    
    return {
      success: true,
      token: token,
      doctors: getActiveDoctors(),
      nurseName: "พยาบาลเวร"
    };
  }
  
  return { success: false, error: "รหัสผ่านไม่ถูกต้อง" };
}

// ---- Verify Session ----
function verifySession(token) {
  if (!token) return false;
  const cached = CacheService.getScriptCache().get("session_" + token);
  return cached === "valid";
}

// ---- Form Status ----
function handleFormStatus(payload) {
  if (!verifySession(payload.token)) {
    return { error: "Session หมดอายุ กรุณา Login ใหม่" };
  }
  
  const status = getQueueStatus();
  // Form เห็น HN เต็ม (เป็นพยาบาลที่ login แล้ว)
  return status;
}

// ---- Submit Case ----
function handleSubmitCase(payload) {
  if (!verifySession(payload.token)) {
    return { error: "Session หมดอายุ กรุณา Login ใหม่" };
  }
  
  const hn = String(payload.hn || "").trim();
  if (!hn) return { error: "กรุณากรอก HN" };
  
  const diagnosis = payload.diagnosis || "";
  const ward = payload.ward || "";
  const note = payload.note || "";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CASES_SHEET);
  
  const nextDoctor = getNextDoctor();
  const totalCases = getTotalCaseCount();
  const queueNumber = totalCases + 1;
  const timestamp = new Date();
  
  // เขียนลง Sheet
  sheet.appendRow([
    timestamp,
    hn,
    diagnosis,
    ward,
    nextDoctor,
    queueNumber,
    note
  ]);
  
  // Format timestamp
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat("dd/MM/yyyy HH:mm");
  highlightRow(sheet, lastRow, nextDoctor);
  
  return {
    success: true,
    hn: hn,
    doctor: nextDoctor,
    queueNumber: queueNumber,
    timestamp: Utilities.formatDate(timestamp, "Asia/Bangkok", "dd/MM/yyyy HH:mm")
  };
}

// ============================================
// ตั้งรหัสผ่านฟอร์มพยาบาล
// ============================================
function setFormPassword() {
  const ui = SpreadsheetApp.getUi();
  const currentPw = getFormPassword();
  
  const response = ui.prompt(
    "🔐 ตั้งรหัสผ่านฟอร์มพยาบาล",
    `รหัสปัจจุบัน: ${currentPw}\n\nใส่รหัสใหม่:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newPw = response.getResponseText().trim();
    if (newPw) {
      PropertiesService.getScriptProperties().setProperty("FORM_PASSWORD", newPw);
      ui.alert(`✅ ตั้งรหัสฟอร์มเรียบร้อย!\n\nรหัสใหม่: ${newPw}\n\n📌 พยาบาลเข้าได้จากหน้าแสดงคิว → กดลิงก์「กรอก Case กลาง」แล้วใส่รหัสนี้`);
    }
  }
}

// ============================================
// 9) สถานะคิวปัจจุบัน
// ============================================
function getQueueStatus() {
  const doctors = getActiveDoctors();
  const counts = getCaseCounts();
  const nextDoctor = getNextDoctor();
  const totalCases = getTotalCaseCount();
  
  // สร้างลำดับคิวถัดไป 5 คน
  const upcomingQueue = getUpcomingQueue(5);
  
  // สรุปจำนวน case ของแต่ละแพทย์
  const doctorStats = doctors.map(doc => ({
    name: doc.name,
    caseCount: counts[doc.name] || 0,
    order: doc.order
  }));
  
  // Case ล่าสุด
  const recentCases = getRecentCases(5);
  
  return {
    timestamp: new Date().toISOString(),
    totalCases: totalCases,
    nextDoctor: nextDoctor,
    upcomingQueue: upcomingQueue,
    doctorStats: doctorStats,
    recentCases: recentCases
  };
}

// ============================================
// 10) คิวถัดไป N คน
// ============================================
function getUpcomingQueue(n) {
  const doctors = getActiveDoctors();
  const counts = getCaseCounts();
  
  // Clone counts
  const simCounts = {};
  doctors.forEach(doc => {
    simCounts[doc.name] = counts[doc.name] || 0;
  });
  
  const queue = [];
  for (let i = 0; i < n; i++) {
    // หาคนที่ case น้อยที่สุด
    let minCount = Infinity;
    doctors.forEach(doc => {
      if (simCounts[doc.name] < minCount) minCount = simCounts[doc.name];
    });
    
    for (const doc of doctors) {
      if (simCounts[doc.name] === minCount) {
        queue.push({
          position: i + 1,
          doctor: doc.name
        });
        simCounts[doc.name]++;
        break;
      }
    }
  }
  
  return queue;
}

// ============================================
// 11) Case ล่าสุด
// ============================================
function getRecentCases(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  
  const cases = [];
  for (let i = data.length - 1; i >= 1 && cases.length < limit; i--) {
    if (data[i][1]) { // มี HN
      cases.push({
        timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), "Asia/Bangkok", "dd/MM/yyyy HH:mm") : "",
        hn: String(data[i][1]),
        diagnosis: data[i][2] || "",
        ward: data[i][3] || "",
        doctor: data[i][4] || "",
        queueNumber: data[i][5] || ""
      });
    }
  }
  
  return cases;
}

// ============================================
// 12) ค้นหาด้วย HN
// ============================================
function searchByHN(hn) {
  if (!hn) return { error: "กรุณาระบุ HN" };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).includes(hn)) {
      results.push({
        timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), "Asia/Bangkok", "dd/MM/yyyy HH:mm") : "",
        hn: String(data[i][1]),
        diagnosis: data[i][2] || "",
        ward: data[i][3] || "",
        doctor: data[i][4] || "",
        queueNumber: data[i][5] || ""
      });
    }
  }
  
  return {
    query: hn,
    count: results.length,
    results: results
  };
}

// ============================================
// 13) เมนูใน Google Sheet
// ============================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🏥 ระบบคิว Case")
    .addItem("⚙️ ตั้งค่าเริ่มต้น", "setupSheets")
    .addItem("📊 ดูสถานะคิว", "showQueueDialog")
    .addItem("🔗 ดู Link ฟอร์มพยาบาล", "showFormLink")
    .addSeparator()
    .addItem("🔐 ตั้งรหัสฟอร์มพยาบาล", "setFormPassword")
    .addItem("🔐 ตั้งรหัส Dashboard Token (PDPA)", "setInternalToken")
    .addItem("🗑️ ตั้ง Auto-Cleanup ข้อมูลเก่า (PDPA)", "setupAutoCleanup")
    .addItem("🧹 ลบข้อมูลเก่าทันที", "autoCleanupOldCases")
    .addSeparator()
    .addItem("🔄 รีเซ็ตคิว (ระวัง!)", "confirmReset")
    .addToUi();
}

function showFormLink() {
  const pw = getFormPassword();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Sarabun, sans-serif; padding: 16px;">
      <h3>📋 ฟอร์มพยาบาล</h3>
      <p>ฟอร์มกรอก Case กลางอยู่ที่ <strong>หน้าเว็บแสดงคิว</strong> (คู่กับ queue-website) แล้ว</p>
      <p style="margin-top:12px;">พยาบาลเปิด <strong>หน้าแสดงคิว</strong> → กดลิงก์「กรอก Case กลาง」→ ใส่รหัส → กรอก Case ได้</p>
      <hr style="margin:12px 0">
      <p><strong>รหัสผ่านฟอร์มปัจจุบัน:</strong> <code>${pw}</code></p>
    </div>
  `).setWidth(450).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, "🔗 Link ฟอร์มพยาบาล");
}

// ============================================
// 14) Dialog แสดงสถานะคิว
// ============================================
function showQueueDialog() {
  const status = getQueueStatus();
  
  let html = '<div style="font-family: Sarabun, sans-serif; padding: 16px;">';
  html += '<h2 style="color: #1a73e8;">📋 สถานะคิว Case กลาง</h2>';
  html += `<p><strong>Case ทั้งหมด:</strong> ${status.totalCases} case</p>`;
  html += `<p style="font-size: 18px; color: #d32f2f;"><strong>🔴 คิวถัดไป: ${status.nextDoctor}</strong></p>`;
  
  html += '<h3>📊 สรุปจำนวน Case</h3><table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;">';
  html += '<tr style="background: #4285f4; color: white;"><th>แพทย์</th><th>จำนวน Case</th></tr>';
  status.doctorStats.forEach(doc => {
    const isNext = doc.name === status.nextDoctor;
    const bg = isNext ? '#FFF3E0' : 'white';
    html += `<tr style="background: ${bg};"><td>${doc.name} ${isNext ? '👈 ถัดไป' : ''}</td><td style="text-align: center;">${doc.caseCount}</td></tr>`;
  });
  html += '</table>';
  
  html += '<h3>🔮 ลำดับคิวถัดไป</h3><ol>';
  status.upcomingQueue.forEach(q => {
    html += `<li><strong>${q.doctor}</strong></li>`;
  });
  html += '</ol></div>';
  
  const ui = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(ui, "ระบบคิว Case กลาง");
}

// ============================================
// 15) รีเซ็ตคิว
// ============================================
function confirmReset() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ ยืนยันรีเซ็ตคิว",
    "การรีเซ็ตจะลบข้อมูล Case ทั้งหมดใน Sheet Cases\nคุณแน่ใจหรือไม่?",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CASES_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    ui.alert("✅ รีเซ็ตเรียบร้อย!");
  }
}

// ============================================
// PDPA: ลบข้อมูลเก่าอัตโนมัติ (Data Retention)
// ตั้ง Trigger ให้ทำงานทุกวัน
// ============================================
function autoCleanupOldCases() {
  const RETENTION_DAYS = 30; // เก็บข้อมูลไว้ 30 วัน
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - RETENTION_DAYS);
  
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const timestamp = new Date(data[i][0]);
    if (timestamp < cutoffDate) {
      rowsToDelete.push(i + 1); // 1-indexed
    }
  }
  
  // ลบจากล่างขึ้นบน
  rowsToDelete.forEach(row => sheet.deleteRow(row));
  
  if (rowsToDelete.length > 0) {
    Logger.log(`PDPA Cleanup: ลบข้อมูลเก่า ${rowsToDelete.length} แถว (เกิน ${RETENTION_DAYS} วัน)`);
  }
}

// ตั้ง Trigger ลบข้อมูลเก่าอัตโนมัติทุกวัน
function setupAutoCleanup() {
  // ลบ trigger เดิมก่อน
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoCleanupOldCases") {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  ScriptApp.newTrigger("autoCleanupOldCases")
    .timeBased()
    .everyDays(1)
    .atHour(2) // ทำงานตี 2
    .create();
  
  SpreadsheetApp.getUi().alert("✅ ตั้ง Auto-Cleanup เรียบร้อย! จะลบข้อมูลเก่ากว่า 30 วันทุกคืน");
}

// ============================================
// PDPA: ตั้งค่า Internal Token
// ============================================
function setInternalToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "🔐 ตั้งรหัส Internal Token",
    "รหัสนี้ใช้เข้าถึง HN เต็มผ่าน Web App\n(สำหรับพยาบาลเท่านั้น)\n\nกรุณาตั้งรหัส:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token) {
      PropertiesService.getScriptProperties().setProperty("INTERNAL_TOKEN", token);
      ui.alert(`✅ ตั้งรหัสเรียบร้อย!\n\nเวลาเข้า Website ในโหมดพยาบาล ให้ใส่รหัสนี้`);
    }
  }
}

// ============================================
// 16) Manual assign (กรณีต้องการระบุแพทย์เอง)
// ============================================
function manualAssign() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet.getName() !== CASES_SHEET) {
    ui.alert("กรุณาเลือก Sheet 'Cases' ก่อนค่ะ");
    return;
  }
  
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert("กรุณาเลือกแถวข้อมูล (ไม่ใช่ Header)");
    return;
  }
  
  const doctors = getActiveDoctors();
  const names = doctors.map(d => d.name).join(", ");
  
  const response = ui.prompt(
    "ระบุแพทย์",
    `เลือกแพทย์:\n${names}`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    sheet.getRange(row, 5).setValue(response.getResponseText());
  }
}
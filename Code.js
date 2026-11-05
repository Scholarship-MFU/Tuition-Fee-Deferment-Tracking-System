// กำหนดชื่อชีตที่จะใช้เก็บข้อมูล
const SHEET_NAME = "Requests";
const SETTINGS_SHEET = "Settings";

// ฟังก์ชันหลักสำหรับเปิดหน้าเว็บ
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบติดตามเอกสารการผ่อนผันค่าเทอม')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันสำหรับเตรียมโครงสร้าง Google Sheets อัตโนมัติ
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ตรวจสอบและสร้างชีต Requests
  let reqSheet = ss.getSheetByName(SHEET_NAME);
  if (!reqSheet) {
    reqSheet = ss.insertSheet(SHEET_NAME);
    // สร้าง Header
    reqSheet.appendRow(['id', 'docNumber', 'semester', 'academicYear', 'email', 'fullName', 'studentId', 'faculty', 'year', 'requestCount', 'totalFee', 'payAmount', 'remainAmount', 'status', 'submitDate', 'status1Date', 'status2Date', 'statusConfirmedDate', 'status3Date']);
    reqSheet.getRange("A1:S1").setFontWeight("bold").setBackground("#f3f4f6");
    reqSheet.setFrozenRows(1);
  }
  
  // ตรวจสอบและสร้างชีต Settings
  let setSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!setSheet) {
    setSheet = ss.insertSheet(SETTINGS_SHEET);
    setSheet.appendRow(['startDate', 'endDate']);
    setSheet.appendRow(['', '']);
    setSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#f3f4f6");
  }
}

// ฟังก์ชันดึงข้อมูลเริ่มต้น (ส่งให้ Frontend ตอนโหลดหน้าแรก)
function getInitialData() {
  setupDatabase(); // ตรวจสอบชีตก่อนดึงเสมอ
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- ดึงข้อมูลคำร้อง ---
  const reqSheet = ss.getSheetByName(SHEET_NAME);
  const reqData = reqSheet.getDataRange().getDisplayValues();
  const reqHeaders = reqData[0];
  const requests = [];
  
  for (let i = 1; i < reqData.length; i++) {
    let obj = {};
    for (let j = 0; j < reqHeaders.length; j++) {
      obj[reqHeaders[j]] = reqData[i][j];
    }
    requests.push(obj);
  }
  
  // --- ดึงข้อมูลการตั้งค่า ---
  const setSheet = ss.getSheetByName(SETTINGS_SHEET);
  const setData = setSheet.getDataRange().getDisplayValues();
  const settings = {
    startDate: setData[1] ? setData[1][0] : '',
    endDate: setData[1] ? setData[1][1] : ''
  };
  
  return { requests: requests, settings: settings };
}

// ฟังก์ชันรับข้อมูลคำร้องจาก Frontend เพื่ออัปเดตลง Sheets
function syncRequests(requests) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // ลบข้อมูลเก่าทั้งหมด (ยกเว้น Header) แล้วเขียนใหม่เพื่อความรวดเร็วของ SPA
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (!requests || requests.length === 0) return true;
  
  const rows = requests.map(req => {
    return headers.map(header => req[header] !== undefined ? req[header] : '');
  });
  
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return true;
}

// ฟังก์ชันรับข้อมูลการตั้งค่าจาก Frontend
function syncSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET);
  sheet.getRange(2, 1, 1, 2).setValues([[settings.startDate, settings.endDate]]);
  return true;
}

// ฟังก์ชันสำหรับส่งอีเมลแจ้งเตือนเมื่ออนุมัติ (เพิ่มใหม่ส่วนท้ายสุด)
function sendApprovalEmail(recipientEmail) {
  const subject = "เอกสารผ่อนผันฯ ได้รับการอนุมัติเเล้ว (กรุณาเปิดอ่านอีเมล) The documents for Tuition Fee Payment Extension have been approved (Please read the following information carefully)";
  
  // หมายเหตุ: กรุณานำรูปภาพไปฝากบนเว็บไซต์ หรือ Google Drive (แบบแชร์สาธารณะ) 
  // แล้วนำลิงก์ (URL) มาใส่แทนที่ "1xT5n1aOLz4iQyhExLn_i5c-BtNLb95rE" และ "1jR70VOw2g1EQo7PQ3bhvaC2EwDSIoGAv"
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto;">
      <p>เอกสารผ่อนผันฯ ได้รับการอนุมัติเเล้ว นักศึกษาสามารถชำระเงินค่าธรรมเนียมการศึกษาที่ได้ทำเรื่องขอผ่อนผันฯ ไว้ โดยสามารถ ชำระได้ตั้งเเต่บัดนี้เป็นต้นไป โดยเร็วที่สุด</p>
      
      <p>โดยนักศึกษาสามารถชำระเงิน online ในเมนู ชำระเงิน ในระบบ REG ของตนเอง (คลิกปุ่ม "ชำระเงิน" ด้วยนะคะ)<br>
      และหากพบว่า จำนวนเงินที่ได้ทำเรื่องขอชำระไว้ ไม่ตรงกับจำนวนเงินที่ปรากฎในระบบ ให้ติดต่อเพื่อเเก้ไขยอดดังกล่าวที่ เบอร์โทรศัพท์ 053-916186</p>
      
      <p><b>ขั้นตอนการชำระเงิน :</b> <a href="https://finance.mfu.ac.th/fa-user/fa-5929/fa-5932.html">https://finance.mfu.ac.th/fa-user/fa-5929/fa-5932.html</a></p>
      
      <p style="color: #d97706;"><b>หมายเหตุ :</b> ในส่วนที่นักศึกษาได้ยื่นเรื่องทำเอกสารผ่อนผันฯ เป็นในส่วนเฉพาะค่าลงทะเบียนเรียนเหมาจ่าย ไม่รวมกับค่าส่วนต่างอื่นๆ เช่น ค่าหอพัก / ค่าบำรุงฯ / ค่าธรรมเนียมอื่นๆ ไม่สามารถผ่อนผันฯ ได้ ดังนั้น อาจมีส่วนต่างอื่นๆ ที่ต้องชำระด้วยค่ะ</p>
      
      <hr style="border: none; border-top: 1px solid #ddd; margin: 25px 0;">
      
      <p>The documents for Tuition Fee Payment Extension have been approved, students can make the tuition fee payment via <a href="https://portal.mfu.ac.th/login?redirectUrl=https%3A%2F%2Freg.mfu.ac.th%2Fbackend" style="color: #2563eb;">REG System</a>. (in menu payment)</p>
      
      <p><b>The tuition fee payment extension how to pay :</b> <a href="https://finance.mfu.ac.th/fa-user/fa-5929/fa-5932.html">https://finance.mfu.ac.th/fa-user/fa-5929/fa-5932.html</a></p>
      
      <p>If the amount of tuition fee payment does not match the amount in the Reg system, please contact Call. 053-916186 for editing the information.</p>
      
      <br>
      </div>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    return true;
  } catch (error) {
    throw new Error(error.toString());
  }
}
function forceAuth() {
  MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getActiveSpreadsheet();
}

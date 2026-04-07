/**
 * โค้ดสำหรับ Google Apps Script เพื่อใช้เป็น API
 * 
 * วิธีการติดตั้ง:
 * 1. เปิด Google Sheet ที่เก็บข้อมูลของคุณ (https://docs.google.com/spreadsheets/d/1K-uQpZn21dM0YzInjE2Lj-ZSuWOf-a-y4Vr1WMtDjWY/edit)
 * 2. ไปที่เมนู Extensions (ส่วนขยาย) -> Apps Script
 * 3. ลบโค้ดเดิมทั้งหมด แล้วนำโค้ดด้านล่างนี้ไปวาง
 * 4. กดปุ่ม Save (ไอคอนแผ่นดิสก์)
 * 5. กดปุ่ม Deploy (การทำให้ใช้งานได้) ที่มุมขวาบน -> New Deployment (การทำให้ใช้งานได้ใหม่)
 * 6. เลือกประเภท (Select type) เป็น Web App (เว็บแอปพลิเคชัน)
 * 7. ตั้งค่าดังนี้:
 *    - Execute as (ดำเนินการในฐานะ): Me (ฉัน - อีเมลของคุณ)
 *    - Who has access (ผู้มีสิทธิ์เข้าถึง): Anyone (ทุกคน)
 * 8. กด Deploy (ทำให้ใช้งานได้) -> อาจจะต้องกดยืนยันสิทธิ์ (Authorize access)
 * 9. Copy Web App URL ที่ได้ (ขึ้นต้นด้วย https://script.google.com/macros/s/...)
 * 10. นำ URL นั้นไปใส่แทนที่ YOUR_SCRIPT_ID ในบรรทัด APPS_SCRIPT_URL ของไฟล์ .env ในโปรเจค Node.js ของเรา
 */

const SHEET_NAME = 'Job';

function doPost(e) {
  try {
    // อ่านข้อมูล JSON ที่ส่งมาจาก Node.js
    const data = JSON.parse(e.postData.contents);
    const key = data.key;
    const newStatus = data.status;
    const signature = data.signature; // รูปแบบ base64
    const dropoff = data.dropoff;
    const remark = data.remark;

    if (!key) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: 'ไม่ได้ระบุ Key' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: 'ไม่พบชีตชื่อ Job' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0]; // แถวที่ 1 เป็น Header

    // หา index ของคอลัมน์ต่างๆ (0-based indexing)
    const keyIdx = headers.indexOf('Key');
    const statusIdx = headers.indexOf('Status');
    const dropoffIdx = headers.indexOf('Dropoff');
    const dropoffSigIdx = headers.indexOf('Dropoff Signature');
    const remarkIdx = headers.indexOf('Remark');

    if (keyIdx === -1 || statusIdx === -1) {
       return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: 'ไม่พบคอลัมน์ Key หรือ Status ในชีต' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let foundRow = -1;

    // เริ่มหาจากแถวที่ 2 (index 1) ลงไป
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][keyIdx]) === String(key)) {
        foundRow = i + 1; // Google Sheet row เริ่มที่ 1
        break;
      }
    }

    if (foundRow !== -1) {
      // อัปเดตสถานะ
      sheet.getRange(foundRow, statusIdx + 1).setValue(newStatus);
      
      // ถ้าเป็นการจบงาน (รับแล้ว)
      if (newStatus === 'Finished') {
        if (dropoffIdx !== -1 && dropoff) sheet.getRange(foundRow, dropoffIdx + 1).setValue(dropoff);
        if (dropoffSigIdx !== -1 && signature) sheet.getRange(foundRow, dropoffSigIdx + 1).setValue(signature);
        if (remarkIdx !== -1 && remark) sheet.getRange(foundRow, remarkIdx + 1).setValue(remark);
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: 'อัปเดตข้อมูลสำเร็จ' 
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: 'ไม่พบข้อมูล Key ที่ระบุ' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      error: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// เพิ่ม doGet เข้ามาเพื่อป้องกัน error เวลาทดสอบเปิด URL โดยตรงผ่านเบราว์เซอร์
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: "DocDelivery Apps Script is running."
  })).setMimeType(ContentService.MimeType.JSON);
}

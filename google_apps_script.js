/**
 * Google Apps Script API สำหรับ DocDelivery
 * 
 * วิธีการติดตั้ง:
 * 1. เปิด Google Sheet → Extensions → Apps Script
 * 2. ลบโค้ดเดิม แล้ววางโค้ดนี้
 * 3. Deploy → New Deployment → Web App → Execute as Me → Anyone
 * 4. Copy URL ไปใส่ใน .env (APPS_SCRIPT_URL)
 * 
 * ** หลังวางโค้ดใหม่ ต้อง Deploy → New Deployment ทุกครั้ง **
 */

const SHEET_NAME = 'Job';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const key = data.key;
    const newStatus = data.status;
    const signature = data.signature; // base64 image data
    const dropoff = data.dropoff;
    const remark = data.remark;

    if (!key) {
      return jsonResponse({ success: false, error: 'ไม่ได้ระบุ Key' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ success: false, error: 'ไม่พบชีตชื่อ Job' });
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    const keyIdx = headers.indexOf('Key');
    const statusIdx = headers.indexOf('Status');
    const dropoffIdx = headers.indexOf('Dropoff');
    const dropoffSigIdx = headers.indexOf('Dropoff Signature');
    const remarkIdx = headers.indexOf('Remark');
    const txt01Idx = headers.indexOf('Txt_01'); // Column S - Signature URL

    if (keyIdx === -1 || statusIdx === -1) {
      return jsonResponse({ success: false, error: 'ไม่พบคอลัมน์ Key หรือ Status' });
    }

    let foundRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][keyIdx]) === String(key)) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return jsonResponse({ success: false, error: 'ไม่พบข้อมูล Key ที่ระบุ' });
    }

    // Update status
    sheet.getRange(foundRow, statusIdx + 1).setValue(newStatus);

    // If finishing the job
    if (newStatus === 'Finished') {
      if (dropoffIdx !== -1 && dropoff) {
        sheet.getRange(foundRow, dropoffIdx + 1).setValue(dropoff);
      }
      if (remarkIdx !== -1 && remark) {
        sheet.getRange(foundRow, remarkIdx + 1).setValue(remark);
      }
      
      // Save signature to Google Drive and store URL in Txt_01 (column S)
      if (signature && signature.startsWith('data:image')) {
        try {
          const sigUrl = saveSignatureToDrive(key, signature);
          if (txt01Idx !== -1 && sigUrl) {
            sheet.getRange(foundRow, txt01Idx + 1).setValue(sigUrl);
          }
          // Also store raw base64 in Dropoff Signature column for backup
          if (dropoffSigIdx !== -1) {
            sheet.getRange(foundRow, dropoffSigIdx + 1).setValue(sigUrl);
          }
        } catch (sigErr) {
          // If Drive upload fails, store base64 directly
          if (dropoffSigIdx !== -1) {
            sheet.getRange(foundRow, dropoffSigIdx + 1).setValue(signature);
          }
        }
      }
    }

    return jsonResponse({ success: true, message: 'อัปเดตข้อมูลสำเร็จ' });

  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Save base64 signature image to Google Drive and return thumbnail URL
 */
function saveSignatureToDrive(key, base64Data) {
  // Extract base64 content
  const parts = base64Data.split(',');
  const mimeMatch = parts[0].match(/:(.*?);/);
  const mime = mimeMatch ? mimeMatch[1] : 'image/png';
  const bytes = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(bytes, mime, 'sig_' + key + '.png');
  
  // Get or create the signatures folder
  const folderName = 'DocDelivery_Signatures';
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
    // Make folder accessible
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  
  // Create the file
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // Return thumbnail URL (same format as existing data)
  const fileId = file.getId();
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return jsonResponse({
    success: true,
    message: "DocDelivery Apps Script is running."
  });
}

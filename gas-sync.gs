// ════════════════════════════════════════════════
//  CTS Training Tracker — GAS Sync + Drive Upload
//  วิธี redeploy:
//  1. แทนที่ code ทั้งหมดใน Apps Script ด้วยไฟล์นี้
//  2. Deploy > Manage deployments > Edit (ดินสอ)
//     > Version: New version > Deploy
//  URL เดิมใช้ได้ต่อ ไม่ต้องแก้ในแอป
// ════════════════════════════════════════════════

var SPREADSHEET_ID = '1u53rZoP87tO2CiZTD7WugV2SE2zpQn7itPZivnddmUQ';
var SHEET_NAME     = 'ClinicsData';
var DRIVE_FOLDER   = 'CTS Training Files';  // ชื่อโฟลเดอร์หลักใน Drive

// ── GET: โหลดข้อมูลคลินิก
function doGet(e) {
  return _load();
}

// ── POST: บันทึกข้อมูล หรือ อัปโหลดไฟล์
function doPost(e) {
  try {
    var body;
    // ตรวจ e.parameter.payload ก่อน (form POST ส่งมาแบบนี้)
    if (e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      // fetch POST ที่ส่ง JSON ตรงๆ (เช่น upload)
      body = JSON.parse(e.postData.contents);
    } else {
      throw new Error('No payload found in request (postData and parameter both empty)');
    }

    if (!body || typeof body !== 'object') {
      throw new Error('Payload parsed but is not a valid object: ' + typeof body);
    }

    if (body.action === 'upload') {
      return _uploadToDrive(body);
    }
    return _saveData(body);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── บันทึกข้อมูลคลินิกลง Sheets
function _saveData(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange('A1').setValue(body.data);
  sh.getRange('B1').setValue(body.ts);
  sh.getRange('C1').setValue(new Date().toISOString());
  sh.getRange('D1').setValue(body.user || 'unknown');
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', ts: body.ts }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── อัปโหลดไฟล์ขึ้น Google Drive แยกโฟลเดอร์ตามเดือน
function _uploadToDrive(body) {
  if (!body || !body.fileName || !body.fileData) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'ข้อมูลไม่ครบ — ต้องการ: fileName, fileData (และ mimeType, month)'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var fileName   = body.fileName;
  var base64     = body.fileData;
  var mimeType   = body.mimeType || 'application/pdf';
  var monthLabel = body.month || 'ไม่ระบุเดือน';

  // โฟลเดอร์หลัก
  var rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  var root = rootFolders.hasNext()
    ? rootFolders.next()
    : DriveApp.createFolder(DRIVE_FOLDER);

  // โฟลเดอร์ย่อยตามเดือน
  var subFolders = root.getFoldersByName(monthLabel);
  var monthFolder = subFolders.hasNext()
    ? subFolders.next()
    : root.createFolder(monthLabel);

  // สร้างไฟล์
  var bytes = Utilities.base64Decode(base64);
  var blob  = Utilities.newBlob(bytes, mimeType, fileName);
  var file  = monthFolder.createFile(blob);

  // เปิดให้ทุกคนที่มีลิงก์ดูได้
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var viewUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      url:    viewUrl,
      id:     file.getId(),
      name:   fileName
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── โหลดข้อมูลคลินิกจาก Sheets
function _load() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      return ContentService
        .createTextOutput(JSON.stringify({ data: '[]', ts: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var data = sh.getRange('A1').getValue() || '[]';
    var ts   = sh.getRange('B1').getValue() || 0;
    return ContentService
      .createTextOutput(JSON.stringify({ data: String(data), ts: Number(ts) }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message, data: '[]', ts: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

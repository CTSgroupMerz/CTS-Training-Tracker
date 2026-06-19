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

    // ── Login / token (ใหม่ — ไม่กระทบ sync เดิม) ──
    if (body.action === 'login') {
      return handleLogin(body.username, body.password);
    }
    if (body.action === 'verify') {
      var info = verifyToken(body.token);
      return ContentService
        .createTextOutput(JSON.stringify(
          info ? { ok: true, username: info.username, displayName: info.displayName }
               : { ok: false }
        ))
        .setMimeType(ContentService.MimeType.JSON);
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

// ════════════════════════════════════════════════
//  LOGIN (ใหม่) — Users sheet + token ใน ScriptProperties
// ════════════════════════════════════════════════

// ใช้ openById(SPREADSHEET_ID) ให้ตรงกับโค้ดเดิม (standalone script — getActiveSpreadsheet เป็น null)
function ensureUsersSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.getRange('A1:C1').setValues([['username','password','displayName']]);
    var defaults = [
      ['parichat',   'cts1234', 'Parichat'],
      ['kanwara',    'cts1234', 'Kanwara'],
      ['narakamon',  'cts1234', 'Narakamon'],
      ['witchukorn', 'cts1234', 'Witchukorn'],
      ['koollanut',  'cts1234', 'Koollanut'],
      ['pariyachat', 'cts1234', 'Pariyachat'],
      ['pitchaporn', 'cts1234', 'Pitchaporn'],
      ['onkamol',    'cts1234', 'Onkamol']
    ];
    sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }
  return sheet;
}

function handleLogin(username, password) {
  var sheet = ensureUsersSheet();
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim() &&
        String(data[i][1]).trim() === String(password).trim()) {

      var token  = Utilities.getUuid();
      var expiry = new Date().getTime() + (8 * 60 * 60 * 1000); // 8 ชม.
      PropertiesService.getScriptProperties()
        .setProperty('tok_' + token, username + '|' + data[i][2] + '|' + expiry);

      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true, token: token, username: username, displayName: data[i][2]
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'invalid' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyToken(token) {
  if (!token) return null;
  var val = PropertiesService.getScriptProperties().getProperty('tok_' + token);
  if (!val) return null;

  var parts  = val.split('|');
  var expiry = parseInt(parts[2]);
  if (new Date().getTime() > expiry) {
    PropertiesService.getScriptProperties().deleteProperty('tok_' + token);
    return null;
  }
  return { username: parts[0], displayName: parts[1] };
}

// Time-driven trigger (รายวัน เที่ยงคืน): ล้าง token หมดอายุ
function cleanExpiredTokens() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var now   = new Date().getTime();
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('tok_') !== 0) return;
    var expiry = parseInt(props[key].split('|')[2]);
    if (now > expiry) {
      PropertiesService.getScriptProperties().deleteProperty(key);
    }
  });
}

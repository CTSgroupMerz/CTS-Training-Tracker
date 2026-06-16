// ════════════════════════════════════════════════
//  CTS Training Tracker — GAS Sync Backend
//  วิธีใช้:
//  1. เปิด Google Apps Script ที่ deploy ไว้
//  2. แทนที่ code ทั้งหมดด้วยไฟล์นี้
//  3. กด Deploy > Manage deployments > Edit (ดินสอ)
//     > Version: New version > Deploy
//  ไม่ต้องเปลี่ยน URL — URL เดิมใช้ได้ต่อ
// ════════════════════════════════════════════════

var SPREADSHEET_ID = '1u53rZoP87tO2CiZTD7WugV2SE2zpQn7itPZivnddmUQ';
var SHEET_NAME = 'ClinicsData';

function doGet(e) {
  return _load();
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    sh.getRange('A1').setValue(body.data);
    sh.getRange('B1').setValue(body.ts);
    sh.getRange('C1').setValue(new Date().toISOString());
    sh.getRange('D1').setValue(body.user || 'unknown');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', ts: body.ts }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

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

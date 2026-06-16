// ════════════════════════════════════════════════
//  CTS Training Tracker — GAS Sync Backend
//  วิธีใช้:
//  1. สร้าง Google Sheets ใหม่
//  2. ก็อป Spreadsheet ID จาก URL แล้วใส่ด้านล่าง
//  3. Deploy > New deployment > Web App
//     Execute as: Me | Who has access: Anyone
//  4. ก็อป URL ไปใส่ใน index.html บรรทัด GAS_URL
// ════════════════════════════════════════════════

var SPREADSHEET_ID = ''; // ← ใส่ Spreadsheet ID ที่นี่
var SHEET_NAME = 'ClinicsData';

function doGet(e) {
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

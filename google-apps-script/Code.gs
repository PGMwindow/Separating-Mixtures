const RESULTS_SHEET_NAME = 'Results';
const RESULTS_HEADERS = [
  'เวลาที่บันทึก',
  'ชื่อ–นามสกุล',
  'ชั้นเรียน',
  'เลขที่',
  'คะแนน',
  'คะแนนเต็ม',
  'ภารกิจที่ได้ 0 คะแนน',
  'ภารกิจที่ทำครบ'
];

function setupResultsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(RESULTS_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(RESULTS_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RESULTS_HEADERS.length).setValues([RESULTS_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, RESULTS_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#dff2ff');
    sheet.autoResizeColumns(1, RESULTS_HEADERS.length);
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: 'Results recorder is ready.' });
}

function doPost(event) {
  setupResultsSheet();
  const values = event && event.parameter ? event.parameter : {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET_NAME);
    sheet.appendRow([
      values.recordedAt || new Date().toISOString(),
      values.name || '',
      values.classroom || '',
      values.number || '',
      values.score || '',
      values.maxScore || '',
      values.zeroScoreMissions || '',
      values.missionsCompleted || ''
    ]);
    return jsonResponse({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

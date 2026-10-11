// Paste into your Google Sheet: Extensions > Apps Script. See README "Google Sheet (optional)".
// Set TOKEN to any secret string and put the same value in config.json -> google_sheet.token.
const TOKEN = 'change-me';

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.token !== TOKEN) return ContentService.createTextOutput('bad token');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const current = ss.getSheetByName('Current') || ss.insertSheet('Current');
  current.clearContents();
  if (body.current.length) {
    const width = Math.max(...body.current.map(r => r.length));
    const rows = body.current.map(r => r.concat(Array(width - r.length).fill('')));
    current.getRange(1, 1, rows.length, width).setValues(rows);
    current.setFrozenRows(1);
  }

  const history = ss.getSheetByName('History') || ss.insertSheet('History');
  if (history.getLastRow() === 0) history.appendRow(['Time', 'Survey', 'Type', 'Details']);
  if (body.events.length) {
    history.getRange(history.getLastRow() + 1, 1, body.events.length, 4).setValues(body.events);
  }
  return ContentService.createTextOutput('ok');
}

/**
 * Property Intake → Google Sheets receiver.
 *
 * Paste this into the Google Sheet: Extensions → Apps Script → replace Code.gs → Save.
 * Change TOKEN below to any long random string, and put the same string in Render as SHEETS_TOKEN.
 * Then Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone → Deploy.
 * Copy the Web app URL into Render as SHEETS_WEBHOOK_URL.
 *
 * The app sends two kinds of request:
 *   {action:"append", header:[...], row:[...]}           a new property row
 *   {action:"update", file:"x.pdf", column:"...", value}  a reviewer's correction
 */
const TOKEN = 'CHANGE-ME-to-a-long-random-string';
const SHEET_NAME = 'Properties';

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.token !== TOKEN) return out({ ok: false, error: 'bad token' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (body.action === 'append') {
    if (sh.getLastRow() === 0) {
      sh.appendRow(body.header);
      sh.getRange(1, 1, 1, body.header.length).setFontWeight('bold').setBackground('#dde3f3').setWrap(true);
      sh.setFrozenRows(1);
      sh.setFrozenColumns(1);
    }
    sh.appendRow(body.row);
    return out({ ok: true });
  }

  if (body.action === 'update') {
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const col = header.indexOf(body.column) + 1;
    if (col < 1) return out({ ok: false, error: 'no such column' });
    const files = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]);
    for (let i = files.length - 1; i >= 1; i--) {          // newest row for that file
      if (files[i] === body.file) {
        sh.getRange(i + 1, col).setValue(body.value).setBackground('#e5f4e4');  // green = corrected by a person
        return out({ ok: true });
      }
    }
    return out({ ok: false, error: 'file not found' });
  }
  return out({ ok: false, error: 'unknown action' });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

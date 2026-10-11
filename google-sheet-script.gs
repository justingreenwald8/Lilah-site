// Lilah waitlist → Google Sheet
// Paste this into the sheet's Extensions → Apps Script, then Deploy → New deployment → Web app
// (Execute as: Me, Who has access: Anyone). Put the /exec URL in FORM_ENDPOINT in public/index.html.

const SHEET_NAME = 'Waitlist';
const NOTIFY_EMAIL = '';   // optional: put your email here to get an email for every signup

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.company) return ContentService.createTextOutput('ok');   // hidden spam trap was filled in: ignore

  // Keep text as text: stops ZIPs like 02134 losing their 0 and blocks spreadsheet formulas.
  const clean = v => {
    const s = String(v || '').slice(0, 300).trim();
    return s ? "'" + s : '';
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Signed up', 'Email', 'Signing up for', 'ZIP', 'Form', 'Page']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), clean(p.email), clean(p.signing_up_for), clean(p.zip), clean(p.form), clean(p.page)]);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail(NOTIFY_EMAIL, 'New Lilah waitlist signup: ' + (p.email || ''),
      'Email: ' + (p.email || '') + '\nFor: ' + (p.signing_up_for || '') + '\nZIP: ' + (p.zip || ''));
  }
  return ContentService.createTextOutput('ok');
}

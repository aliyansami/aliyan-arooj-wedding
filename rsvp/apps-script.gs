/**
 * RSVP receiver for the Aliyan & Arooj wedding invite.
 *
 * Setup (one time):
 * 1. Open https://sheets.new and name the sheet "Wedding RSVPs".
 * 2. Extensions > Apps Script. Delete what is there and paste this whole file. Save.
 * 3. Deploy > New deployment > type "Web app".
 *    Execute as: Me.  Who has access: Anyone.  Click Deploy and allow access.
 * 4. Copy the Web app URL (ends in /exec) and send it to be added to the invite.
 */
const SHEET_NAME = 'RSVPs';
const HEADERS = ['Submitted', 'Name', 'Phone', 'Attending', 'Guests',
  'Nikkah & Mehndi', 'Baraat', 'Walima', 'Joining from', 'Message'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    const p = (e && e.parameter) || {};
    if (p.website) return ok_(); // spam trap
    const clean = v => {
      const s = String(v || '').slice(0, 500).trim();
      return /^[=+\-@]/.test(s) ? "'" + s : s; // stop formula injection
    };
    sheet.appendRow([
      new Date(),
      clean(p.name),
      p.phone ? "'" + String(p.phone).slice(0, 20) : '',
      clean(p.attending),
      Number(p.guests) || 0,
      clean(p.mehndi),
      clean(p.baraat),
      clean(p.walima),
      clean(p.side),
      clean(p.message)
    ]);
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ok_(); }

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

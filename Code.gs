/**
 * Renzel & Erika wedding website + RSVP → Google Sheet
 * Paste this into Code.gs of the Apps Script project bound to (or with access to) your RSVP sheet.
 */
const SHEET_ID = '1A6Keecu8mdxrsSvfOI3P8zBXKk-IjW4XNutjPhak02A';
const HEADERS = ['Timestamp', 'Name', 'Mobile', 'Email', 'Reply', 'Attending Guest', 'Message'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Renzel & Erika')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called from the RSVP form. Adds a row, or updates the guest's earlier reply (matched by mobile number). */
function submitRsvp(v) {
  v = v || {};
  const name = String(v.name || '').trim();
  const mobile = String(v.mobile || '').trim();
  if (!name || !mobile) throw new Error('Please enter your name and mobile number.');
  if (v.attending !== 'yes' && v.attending !== 'no') throw new Error('Please choose whether you will join us.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    sheet.getRange('C:C').setNumberFormat('@'); // keep the leading 0 in mobile numbers

    const clean = s => String(s || '').trim().slice(0, 1000).replace(/^[=+\-@]/, "'$&");
    const row = [
      new Date(),
      clean(name),
      clean(mobile),
      clean(v.email),
      v.attending === 'yes' ? 'Joyfully accepts' : 'Regretfully declines',
      v.attending === 'yes' ? clean(v.guest || name) : '',
      clean(v.message)
    ];

    const key = mobile.replace(/\D/g, '');
    const last = sheet.getLastRow();
    if (last > 1) {
      const mobiles = sheet.getRange(2, 3, last - 1, 1).getValues();
      for (let i = 0; i < mobiles.length; i++) {
        if (String(mobiles[i][0]).replace(/\D/g, '') === key) {
          sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
          return { updated: true };
        }
      }
    }
    sheet.appendRow(row);
    return { updated: false };
  } finally {
    lock.releaseLock();
  }
}

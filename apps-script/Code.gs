/**
 * JIN YUZE RSVP backend.
 * 1. Open your Google Sheet > Extensions > Apps Script, paste this file.
 * 2. Run setup() once (creates the RSVP and SUMMARY tabs).
 * 3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 * 4. Copy the /exec URL into SCRIPT_URL in index.html.
 */

const RSVP_SHEET = 'RSVP';
const SUMMARY_SHEET = 'SUMMARY';
const HEADERS = ['Name', 'Adult_Guest', 'Child_Guest', 'Total_Guest', 'Attendance', 'Message', 'Submitted_At'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rsvp = ss.getSheetByName(RSVP_SHEET) || ss.insertSheet(RSVP_SHEET);
  rsvp.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  rsvp.setFrozenRows(1);

  const summary = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  summary.clear();
  summary.getRange('A1:B7').setValues([
    ['Metric', 'Value'],
    ['Total RSVP Responses', '=COUNTA(RSVP!A2:A)'],
    ['Total Attending Responses', '=COUNTIF(RSVP!E2:E,"ATTENDING")'],
    ['Total Not Attending', '=COUNTIF(RSVP!E2:E,"NOT ATTENDING")'],
    ['Total Adults Attending', '=SUMIF(RSVP!E2:E,"ATTENDING",RSVP!B2:B)'],
    ['Total Children Attending', '=SUMIF(RSVP!E2:E,"ATTENDING",RSVP!C2:C)'],
    ['Total Confirmed Guests', '=SUMIF(RSVP!E2:E,"ATTENDING",RSVP!D2:D)'],
  ]);
  summary.getRange('A1:B1').setFontWeight('bold');
  summary.getRange('A7:B7').setFontWeight('bold');
  summary.autoResizeColumn(1);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const name = String(data.name || '').trim().slice(0, 100);
    const attendance = data.attendance;
    const message = String(data.message || '').trim().slice(0, 1000);
    let adult = Number(data.adult_guest);
    let child = Number(data.child_guest);

    if (!name) return json({ ok: false, error: 'Please enter your name.' });
    if (attendance !== 'ATTENDING' && attendance !== 'NOT ATTENDING') {
      return json({ ok: false, error: 'Please choose your attendance.' });
    }
    if (attendance === 'NOT ATTENDING') {
      adult = 0;
      child = 0;
    } else {
      if (!Number.isInteger(adult) || adult < 0 || !Number.isInteger(child) || child < 0) {
        return json({ ok: false, error: 'Guest counts must be whole numbers, 0 or more.' });
      }
      if (adult + child < 1) return json({ ok: false, error: 'Please add at least one guest.' });
    }
    // Never trust the frontend total.
    const total = adult + child;

    const tz = Session.getScriptTimeZone();
    const submittedAt = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd hh:mm a');

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET)
        .appendRow([safe(name), adult, child, total, attendance, safe(message), submittedAt]);
    } finally {
      lock.releaseLock();
    }

    return json({ ok: true, name: name, total_guest: total, attendance: attendance });
  } catch (err) {
    return json({ ok: false, error: 'Something went wrong. Please try again.' });
  }
}

// Stop guest text like "=HYPERLINK(...)" from running as a formula in the sheet.
function safe(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

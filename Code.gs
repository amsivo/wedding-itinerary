/**
 * Family Itinerary: Google Sheet backend
 * Paste this into Extensions → Apps Script inside your Google Sheet.
 * See SETUP.md for the full steps.
 */

// The passcode family members type when adding an event from the website.
// Change it before deploying, and only share it with the people who add events.
const PASSCODE = 'change-me';

const EVENTS_SHEET = 'Events';
const PEOPLE_SHEET = 'People';
const HEADERS = ['id', 'date', 'start', 'end', 'title', 'location', 'address', 'people', 'dress', 'notes', 'added_by', 'added_at'];

/** Run this once from the Apps Script editor. It creates the Events and People tabs. */
function setup() {
  const ss = SpreadsheetApp.getActive();

  let events = ss.getSheetByName(EVENTS_SHEET);
  if (!events) events = ss.insertSheet(EVENTS_SHEET, 0);
  if (events.getLastRow() === 0) {
    events.appendRow(HEADERS);
    events.setFrozenRows(1);
    events.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#ECE4D5');
  }
  // Keep start/end as plain text ("6:30 pm", "18:30") so nothing gets reformatted.
  events.getRange('C:D').setNumberFormat('@');
  // Date column gets a date picker (double-click a cell).
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true)
    .setHelpText('Double-click to pick a date').build();
  events.getRange('B2:B').setDataValidation(dateRule).setNumberFormat('ddd d mmm yyyy');

  let people = ss.getSheetByName(PEOPLE_SHEET);
  if (!people) people = ss.insertSheet(PEOPLE_SHEET, 1);
  if (people.getLastRow() === 0) {
    people.appendRow(['name']);
    people.setFrozenRows(1);
    people.getRange(1, 1).setFontWeight('bold').setBackground('#ECE4D5');
  }

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(blank);
}

/** Website reads the itinerary. */
function doGet() {
  try {
    return json({ ok: true, people: readPeople(), events: readEvents() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Website adds an event. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (String(body.passcode || '').trim() !== PASSCODE) return json({ ok: false, error: 'passcode' });

    const ev = body.event || {};
    const title = text(ev.title, 120);
    const date = String(ev.date || '').trim();
    if (!title || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return json({ ok: false, error: 'missing title or date' });

    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS_SHEET);
    if (!sheet) return json({ ok: false, error: 'Run setup() first' });

    const id = Utilities.getUuid().slice(0, 8);
    const [y, m, d] = date.split('-').map(Number);
    const row = {
      id: id,
      date: new Date(y, m - 1, d, 12),          // noon avoids any timezone date shift
      start: text(ev.start, 12),
      end: text(ev.end, 12),
      title: title,
      location: text(ev.location, 120),
      address: text(ev.address, 200),
      people: text(ev.people, 300) || 'Everyone',
      dress: text(ev.dress, 60),
      notes: text(ev.notes, 600),
      added_by: text(ev.addedBy, 60),
      added_at: new Date()
    };
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
    sheet.appendRow(header.map(h => (h in row ? row[h] : '')));
    return json({ ok: true, id: id });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function readEvents() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(EVENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  const values = sheet.getDataRange().getValues();
  const header = values.shift().map(h => String(h).trim().toLowerCase());
  const col = name => header.indexOf(name);

  return values.map((r, i) => {
    const get = name => (col(name) < 0 ? '' : r[col(name)]);
    const asDate = v => (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim());
    const asTime = v => (v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm') : String(v || '').trim());
    return {
      id: String(get('id') || 'row' + (i + 2)),
      date: asDate(get('date')),
      start: asTime(get('start')),
      end: asTime(get('end')),
      title: String(get('title') || '').trim(),
      location: String(get('location') || '').trim(),
      address: String(get('address') || '').trim(),
      people: String(get('people') || '').trim(),
      dress: String(get('dress') || '').trim(),
      notes: String(get('notes') || '').trim()
    };
  }).filter(ev => ev.title);
}

function readPeople() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(PEOPLE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(Boolean);
}

// Trim, cap length, and stop anything that looks like a spreadsheet formula.
function text(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

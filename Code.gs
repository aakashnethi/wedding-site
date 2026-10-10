// Google Apps Script backend for the wedding RSVP page.
//
// Setup (free, uses your Google account):
// 1. Create a Google Sheet. Rename the first tab to "RSVPs".
// 2. Extensions > Apps Script. Replace the default code with this file.
// 3. Deploy > New deployment > type: Web app.
//    Execute as: Me.  Who has access: Anyone.
// 4. Copy the web app URL into CONFIG.endpoint in index.html.
//
// The sheet holds everything, including emails and notes. Only the names
// of guests who chose to be public are ever sent to the website.

const SHEET_NAME = "RSVPs";
const HEADERS = ["Timestamp", "Name", "Email", "Attending", "Events", "Guests", "Note", "Show name"];
const EVENTS = ["Engagement", "Haldi & Sangeet", "Wedding"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  const name = String(data.name || "").trim().slice(0, 80);
  if (!name) return json_({ ok: false, error: "Name required" });

  const attending = data.attending === "yes" ? "Yes" : "No";
  const guests = Math.min(Math.max(parseInt(data.guests, 10) || 1, 1), 10);
  const events = (Array.isArray(data.events) ? data.events : [])
    .filter(ev => EVENTS.includes(ev));

  getSheet_().appendRow([
    new Date(),
    name,
    String(data.email || "").slice(0, 120),
    attending,
    attending === "Yes" ? events.join(", ") : "",
    attending === "Yes" ? guests : 0,
    String(data.note || "").slice(0, 500),
    data.showName ? "Yes" : "No",
  ]);

  return json_({ ok: true });
}

// Public data only: headcounts per event (anonymous guests included in the
// counts), plus the named guests who chose to show their name.
function doGet() {
  const rows = getSheet_().getDataRange().getValues().slice(1);

  const perEvent = {};
  EVENTS.forEach(ev => perEvent[ev] = 0);
  let totalGuests = 0;
  let anonymousCount = 0;
  const named = [];

  rows.forEach(([, name, , attending, events, guests, , showName]) => {
    if (attending !== "Yes") return;
    const count = Number(guests) || 0;
    const attendingEvents = String(events).split(", ").filter(ev => EVENTS.includes(ev));
    totalGuests += count;
    attendingEvents.forEach(ev => perEvent[ev] += count);
    if (showName === "Yes") {
      named.push({ name, events: attendingEvents });
    } else {
      anonymousCount += count;
    }
  });

  return json_({ totalGuests, anonymousCount, perEvent, named });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

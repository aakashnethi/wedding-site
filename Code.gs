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

  const rsvp = {
    name,
    email: String(data.email || "").trim().slice(0, 120),
    attending,
    events: attending === "Yes" ? events.join(", ") : "",
    guests: attending === "Yes" ? guests : 0,
    note: String(data.note || "").slice(0, 500),
    showName: data.showName ? "Yes" : "No",
  };

  getSheet_().appendRow([
    new Date(),
    rsvp.name,
    rsvp.email,
    rsvp.attending,
    rsvp.events,
    rsvp.guests,
    rsvp.note,
    rsvp.showName,
  ]);

  // A failed notification should never lose the RSVP, so errors are only logged.
  try {
    notify_(rsvp);
  } catch (err) {
    console.error("Notification failed: " + err);
  }

  return json_({ ok: true });
}

// Sends the host an email and optional chat message for every RSVP, and a
// confirmation to the guest when they gave a valid email address.
//
// Set these in Apps Script: Project Settings → Script Properties
//   HOST_EMAILS  comma-separated addresses, e.g. "you@gmail.com,partner@gmail.com"
//   WEBHOOK_URL  optional Slack or Discord incoming webhook URL
function notify_(rsvp) {
  const props = PropertiesService.getScriptProperties();
  const hostEmails = props.getProperty("HOST_EMAILS");
  const webhook = props.getProperty("WEBHOOK_URL");

  const attendingText = rsvp.attending === "Yes"
    ? `${rsvp.guests} guest(s) attending: ${rsvp.events || "no events selected"}`
    : "Not attending";
  const details = [
    `Name: ${rsvp.name}`,
    `Response: ${attendingText}`,
    rsvp.email ? `Email: ${rsvp.email}` : null,
    rsvp.note ? `Note: ${rsvp.note}` : null,
    `Show name on guest list: ${rsvp.showName}`,
  ].filter(Boolean).join("\n");

  if (hostEmails) {
    MailApp.sendEmail({
      to: hostEmails,
      subject: `RSVP: ${rsvp.name} (${rsvp.attending === "Yes" ? "attending" : "declined"})`,
      body: `A new RSVP was received.\n\n${details}`,
    });
  }

  if (webhook) {
    const text = `New RSVP\n${details}`;
    UrlFetchApp.fetch(webhook, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text, content: text }), // Slack uses "text", Discord uses "content"
      muteHttpExceptions: true,
    });
  }

  if (rsvp.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(rsvp.email)) {
    const body = rsvp.attending === "Yes"
      ? `Dear ${rsvp.name},\n\nThank you for your RSVP. We look forward to celebrating with you on:\n` +
        `${rsvp.events || "the events you selected"}\n\n` +
        `Guests: ${rsvp.guests}\n\nWith love,\nRishika & Akshay`
      : `Dear ${rsvp.name},\n\nThank you for letting us know. We'll miss you, and we hope to see you soon.\n\n` +
        `With love,\nRishika & Akshay`;
    MailApp.sendEmail({
      to: rsvp.email,
      subject: "Thank you for your RSVP — Rishika & Akshay",
      body,
    });
  }
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

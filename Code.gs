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
// New columns are added at the end so existing rows and the guest-list reader stay aligned.
const HEADERS = ["Timestamp", "Name", "Email", "Attending", "Events", "Guests", "Note", "Show name", "Phone", "Text OK", "Meal"];
const EVENTS = ["Engagement", "Haldi & Sangeet", "Wedding"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

const PHOTOS_SHEET = "Photos";
// Photo ID links each sheet row to its Drive file (it's also in the file name and description).
const PHOTO_HEADERS = ["Timestamp", "Event", "Uploader", "Hashtag", "Caption", "File ID", "Show on site", "Photo ID", "Sharing"];

function getPhotosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PHOTOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PHOTOS_SHEET);
    sheet.appendRow(PHOTO_HEADERS);
  }
  return sheet;
}

// Hashtags are stored without "#" or spaces. "#RiCh" and "rich" group together.
function normalizeTag_(raw) {
  return String(raw || "").replace(/^#+/, "").replace(/[^A-Za-z0-9]/g, "").slice(0, 30);
}

// Photos the couple hasn't hidden. Drive thumbnail links work for files shared "anyone with the link".
function publicPhotos_() {
  const rows = getPhotosSheet_().getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[5] && r[6] !== "No")
    .slice(-200)
    .map(([, event, uploader, hashtag, caption, fileId]) => ({
      event,
      uploader,
      hashtag,
      caption,
      thumb: `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`,
      url: `https://drive.google.com/thumbnail?id=${fileId}&sz=w1600`,
    }))
    .reverse();
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.type === "photo") return savePhoto_(data);

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
    phone: normalizePhone_(data.phone),
    textOk: data.textOk && normalizePhone_(data.phone) ? "Yes" : "No",
    meal: data.meal === "Non-Veg" ? "Non-Veg" : "Veg",
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
    rsvp.phone,
    rsvp.textOk,
    rsvp.meal,
  ]);

  // A failed notification should never lose the RSVP, so errors are only logged.
  try {
    notify_(rsvp);
  } catch (err) {
    console.error("Notification failed: " + err);
  }

  return json_({ ok: true });
}

// Saves one guest photo as a JPEG in the Drive folder named by PHOTOS_FOLDER_ID.
// Photos are capped at 5 MB, and the event name goes into the file name.
function savePhoto_(data) {
  const folderId = PropertiesService.getScriptProperties().getProperty("PHOTOS_FOLDER_ID");
  if (!folderId) return json_({ ok: false, error: "Photo folder not configured" });

  const bytes = Utilities.base64Decode(String(data.data || ""));
  if (!bytes.length || bytes.length > 5 * 1024 * 1024) {
    return json_({ ok: false, error: "Photo missing or too large" });
  }

  const photoId = Utilities.getUuid().slice(0, 8);
  const event = EVENTS.includes(data.event) ? data.event : "Other";
  const uploader = String(data.name || "guest").replace(/[^\w ]/g, "").trim().slice(0, 40) || "guest";
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const blob = Utilities.newBlob(bytes, "image/jpeg", `${photoId} - ${event} - ${uploader} - ${stamp}.jpg`);

  const hashtag = normalizeTag_(data.hashtag);
  const caption = String(data.caption || "").trim().slice(0, 140);

  // The sheet row is written as soon as the file exists, so a sharing problem can't lose the photo.
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  file.setDescription(`Photo ID: ${photoId}`);
  let sharing = "Yes";
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    sharing = "Not shared";
  }
  getPhotosSheet_().appendRow([
    new Date(), event, uploader, hashtag, caption, file.getId(), "Yes", photoId, sharing,
  ]);

  return json_({ ok: true });
}

// Keeps digits and a leading "+". A 10-digit number is treated as a US number.
function normalizePhone_(raw) {
  const cleaned = String(raw || "").replace(/[^\d+]/g, "").slice(0, 20);
  if (/^\d{10}$/.test(cleaned)) return "+1" + cleaned;
  if (/^\+\d{8,15}$/.test(cleaned)) return cleaned;
  return "";
}

// Sends one SMS through Twilio. Does nothing if Twilio isn't configured.
function sendSms_(to, body) {
  const props = PropertiesService.getScriptProperties();
  const sid = props.getProperty("TWILIO_SID");
  const token = props.getProperty("TWILIO_TOKEN");
  const from = props.getProperty("TWILIO_FROM");
  if (!sid || !token || !from || !to) return;

  UrlFetchApp.fetch(`https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`, {
    method: "post",
    payload: { To: to, From: from, Body: body },
    headers: { Authorization: "Basic " + Utilities.base64Encode(sid + ":" + token) },
    muteHttpExceptions: true,
  });
}

// Sends the host an email and optional chat message for every RSVP, and a
// confirmation to the guest when they gave a valid email address.
//
// Set these in Apps Script: Project Settings → Script Properties
//   HOST_EMAILS    comma-separated addresses, e.g. "you@gmail.com,partner@gmail.com"
//   WEBHOOK_URL    optional Slack or Discord incoming webhook URL
//   HOST_PHONES    optional comma-separated numbers to text, e.g. "+15551234567"
//   TWILIO_SID, TWILIO_TOKEN, TWILIO_FROM   optional, to enable text messages
function notify_(rsvp) {
  const props = PropertiesService.getScriptProperties();
  const hostEmails = props.getProperty("HOST_EMAILS");
  const webhook = props.getProperty("WEBHOOK_URL");
  const hostPhones = props.getProperty("HOST_PHONES");

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

  const summaryText = `RSVP from ${rsvp.name}: ${rsvp.attending === "Yes" ? rsvp.guests + " attending (" + rsvp.events + ")" : "declined"}`;
  if (hostPhones) {
    hostPhones.split(",").map(p => normalizePhone_(p)).filter(Boolean)
      .forEach(to => sendSms_(to, summaryText));
  }

  // Texts only go to guests who ticked "OK to send me text messages".
  if (rsvp.textOk === "Yes") {
    const body = rsvp.attending === "Yes"
      ? `Thank you, ${rsvp.name}! We're excited to celebrate with you. — Rishika & Akshay`
      : `Thank you for letting us know, ${rsvp.name}. We'll miss you. — Rishika & Akshay`;
    sendSms_(rsvp.phone, body);
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

  return json_({ totalGuests, anonymousCount, perEvent, named, photos: publicPhotos_() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run once from the Apps Script editor to confirm the Drive permission is granted.
function authorizeDrive() {
  DriveApp.getRootFolder();
}

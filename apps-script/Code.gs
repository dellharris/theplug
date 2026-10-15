var TAB = "RSVPs";
var EVENT_TITLE = "THE PLUG Info Session (IDU House)";
var EVENT_START = "20261015T010000Z";
var EVENT_END = "20261015T020000Z";
var EVENT_PLACE = "IDU House Discord";
var EVENT_DETAILS = "The info session for THE PLUG, the new IDU House creator network. Wednesday at 9 PM ET on the IDU House Discord.";

var SHEET_ID = "1-isMv3HltjEWozb6YVZ6Q8A2w81CVbSY9oUZG0n7ef0";

function ss_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function tab_() {
  var ss = ss_();
  var sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) sh.appendRow(["Timestamp", "Name", "Email", "Status"]);
  return sh;
}

function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  var p = e.parameter || {};
  if (p.form === "feedback") return saveFeedback_(p);
  var name = String(p.name || "").trim().slice(0, 100);
  var email = String(p.email || "").trim().toLowerCase().slice(0, 200);
  if (p.website || !name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
    return ContentService.createTextOutput("ignored");
  }
  var sh = tab_();
  var last = sh.getLastRow();
  var existing = last > 1 ? sh.getRange(2, 3, last - 1, 1).getValues().flat() : [];
  if (existing.indexOf(email) === -1) sh.appendRow([new Date(), safe_(name), email, "pending"]);
  return ContentService.createTextOutput("ok");
}

function doGet() {
  return ContentService.createTextOutput("ok");
}

function saveFeedback_(p) {
  if (p.website) return ContentService.createTextOutput("ignored");
  var clip = function (v, n) { return safe_(String(v || "").trim().slice(0, n)); };
  var ss = ss_();
  var sh = ss.getSheetByName("Feedback") || ss.insertSheet("Feedback");
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Timestamp", "Name", "Email", "I am a", "Clarity (1-5)", "Excitement (1-5)", "Stood out", "Questions", "First campaign or content", "Beta access", "Attended"]);
  }
  sh.appendRow([new Date(), clip(p.name, 100), clip(p.email, 200), clip(p.role, 200), clip(p.clarity, 5), clip(p.excitement, 5), clip(p.standout, 300), clip(p.questions, 2000), clip(p.first, 500), clip(p.beta, 50), clip(p.attended, 20)]);
  return ContentService.createTextOutput("ok");
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("The Plug").addItem("Send confirmations to pending RSVPs", "sendConfirmations").addToUi();
}

function ics_() {
  return [
    "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//IDU House//The Plug//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
    "BEGIN:VEVENT", "UID:plug-cohort-unveil-20261014@iduhouse.com", "DTSTAMP:20261012T000000Z",
    "DTSTART:" + EVENT_START, "DTEND:" + EVENT_END,
    "SUMMARY:" + EVENT_TITLE, "DESCRIPTION:" + EVENT_DETAILS, "LOCATION:" + EVENT_PLACE,
    "END:VEVENT", "END:VCALENDAR"
  ].join("\r\n") + "\r\n";
}

function googleLink_() {
  return "https://calendar.google.com/calendar/render?action=TEMPLATE" +
    "&text=" + encodeURIComponent(EVENT_TITLE) +
    "&dates=" + EVENT_START + "/" + EVENT_END +
    "&details=" + encodeURIComponent(EVENT_DETAILS) +
    "&location=" + encodeURIComponent(EVENT_PLACE);
}

function sendConfirmations() {
  var sh = tab_();
  var last = sh.getLastRow();
  if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();
  var ics = Utilities.newBlob(ics_(), "text/calendar", "plug-info-session.ics");
  var sent = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][3] !== "pending") continue;
    var name = String(rows[i][1]).replace(/^'/, "");
    var link = googleLink_();
    var text = "Hi " + name + ",\n\nYou are confirmed for THE PLUG info session, Wednesday October 14 at 9 PM ET on the IDU House Discord.\n\n" +
      "Add it to your calendar:\nGoogle Calendar: " + link + "\nApple Calendar or Outlook: open the attached plug-info-session.ics file.\n\nSee you there,\nIDU House";
    var safeName = name.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    var html = "<p>Hi " + safeName + ",</p><p>You are confirmed for <strong>THE PLUG info session</strong>, Wednesday October 14 at 9 PM ET on the IDU House Discord.</p>" +
      "<p>Add it to your calendar:<br><a href=\"" + link + "\">Add to Google Calendar</a><br>Apple Calendar or Outlook: open the attached <em>plug-info-session.ics</em> file.</p><p>See you there,<br>IDU House</p>";
    MailApp.sendEmail({ to: rows[i][2], subject: "You are confirmed: THE PLUG info session", body: text, htmlBody: html, name: "IDU House", attachments: [ics] });
    sh.getRange(i + 2, 4).setValue("sent");
    sent++;
  }
  ss_().toast("Sent " + sent + " confirmation(s).");
}

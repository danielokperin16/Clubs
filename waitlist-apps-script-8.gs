/**
 * Clubs waitlist: saves signups to this Google Sheet, emails YOU every time someone joins,
 * emails each signup a confirmation, and sends you a daily summary.
 * Free: runs on your Google account. See waitlist-setup-guide.md for the steps.
 */

// 1) The email that receives a message for every signup and the daily summary.
const NOTIFY_EMAIL = 'danielokperin@gmail.com';

const SHEET_NAME = 'Waitlist';
// Name is last so older rows and the email column stay where they were.
const HEADERS = ['Time', 'Email', 'Role', 'Extra', 'Club clicked', 'Source', 'Consent', 'Name'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (current.join('|') !== HEADERS.join('|')) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Trim, remove line breaks, cap the length, and stop spreadsheet formulas
// (text starting with = + - @) from running.
function clean_(v, n) {
  let t = String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, n);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

// Receives each signup from the landing page.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const email = String(p.email || '').trim().toLowerCase();

    // Basic checks: valid-looking email and consent given.
    if (!/^[A-Za-z0-9_][^\s@]*@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: 'invalid email' });
    if (p.consent !== 'yes') return json_({ ok: false, error: 'no consent' });

    const sheet = getSheet_();
    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2, 2, last - 1, 1).getValues().flat() : [];
    if (existing.indexOf(email) !== -1) return json_({ ok: true, duplicate: true });

    const signup = {
      name: clean_(p.name, 80),
      email: email,
      role: clean_(p.role, 30),
      extra: clean_(p.extra, 120),
      club: clean_(p.club, 120),
      source: clean_(p.source, 120)
    };
    sheet.appendRow([new Date(), signup.email, signup.role, signup.extra, signup.club, signup.source, 'yes', signup.name]);

    const total = sheet.getLastRow() - 1;
    sendOwnerAlert_(signup, total);   // your email comes first
    sendConfirmation_(signup);        // then the signup's confirmation
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Emails YOU as soon as someone joins.
function sendOwnerAlert_(s, total) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) return;
    const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(
      to,
      'New Clubs signup: ' + (s.name || s.email) + ' (' + (s.role || 'unknown') + '), #' + total,
      'Someone just joined the Clubs waitlist.\n\n' +
      'Name: ' + (s.name || '(none)') + '\n' +
      'Email: ' + s.email + '\n' +
      'Joining as: ' + (s.role || 'unknown') + '\n' +
      (s.extra ? 'Details: ' + s.extra + '\n' : '') +
      (s.club ? 'Clicked club: ' + s.club + '\n' : '') +
      'Came from: ' + (s.source || 'unknown') + '\n\n' +
      'Total on the waitlist: ' + total,
      { name: 'Clubs waitlist' }
    );
  } catch (err) {
    // The signup is already saved, so a failed alert never loses it.
  }
}

// Sends the signup a short confirmation. Skips quietly if today's email limit is nearly used up.
function sendConfirmation_(s) {
  try {
    if (MailApp.getRemainingDailyQuota() < 10) return; // keep emails for your alerts and daily summary
    const hello = s.name ? 'Hi ' + s.name.split(' ')[0] + ',' : 'Hi,';
    MailApp.sendEmail(
      s.email,
      "You're on the Clubs waitlist",
      hello + '\n\n' +
      "Thanks for joining the Clubs waitlist. We'll email you when Clubs launches.\n\n" +
      "If you didn't sign up, or you'd like to be removed, just reply to this email.\n\n" +
      'The Clubs team',
      { name: 'Clubs' }
    );
  } catch (err) {
    // The signup is already saved, so a failed confirmation never loses it.
  }
}

// Optional: opening the web app link in a browser shows that it is running.
function doGet() {
  return json_({ ok: true, service: 'clubs waitlist' });
}

// Emails you the total, a role breakdown, and who joined in the last 24 hours.
function sendDailyDigest() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  const total = Math.max(last - 1, 0);
  const rows = total ? sheet.getRange(2, 1, total, HEADERS.length).getValues() : [];

  const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const recent = rows.filter(r => r[0] instanceof Date && r[0] >= since);

  const byRole = {};
  rows.forEach(r => {
    const k = r[2] || 'unknown';
    byRole[k] = (byRole[k] || 0) + 1;
  });
  const roleLines = Object.keys(byRole).map(k => '  ' + k + ': ' + byRole[k]).join('\n') || '  none yet';

  const MAX = 50;
  const newLines = recent.slice(0, MAX).map(r =>
    '  ' + (r[7] || '(no name)') + ', ' + r[1] + ', ' + (r[2] || 'unknown') + (r[3] ? ', ' + r[3] : '')
  ).join('\n') || '  nobody new';
  const more = recent.length > MAX ? '\n  ...and ' + (recent.length - MAX) + ' more in the sheet' : '';

  MailApp.sendEmail(
    NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(),
    'Clubs waitlist: ' + total + ' total (' + recent.length + ' new in 24h)',
    'Total signups: ' + total + '\n' +
    'New in the last 24 hours: ' + recent.length + '\n\n' +
    'By role:\n' + roleLines + '\n\n' +
    'New signups:\n' + newLines + more
  );
}

// Run this ONCE by hand to schedule the daily email at about 8am (your account's time zone).
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyDigest').timeBased().everyDays(1).atHour(8).create();
}

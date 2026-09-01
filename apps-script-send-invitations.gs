/**
 * Singh Foundation — Diwali 2026 invitation sender
 * ------------------------------------------------
 * Sends the invitation as REAL HTML from your own Gmail account, so the
 * black background survives. Pasting into the Gmail compose window strips
 * background colours; this does not.
 *
 * SETUP (once):
 *  1. Make a Google Sheet. Row 1 = headers: A1 "Email", B1 "Status".
 *     Put one guest email per row from A2 down. Leave column B empty.
 *  2. In that Sheet: Extensions > Apps Script.
 *  3. Paste this file over the Code.gs that opens.
 *  4. Click the + next to Files > HTML, name it exactly:  invitation
 *     Delete everything in it, then paste in the whole contents of
 *     diwali-2026-EMAIL-VERSION.html. Save.
 *  5. Run sendTestToMe first. Google will ask you to authorise it — that
 *     is normal, it is your own script sending from your own account.
 *  6. When the test looks right, run sendInvitations.
 *
 * Re-running sendInvitations is safe: rows already marked SENT are skipped,
 * so nobody is invited twice if you stop and start again.
 */

const SUBJECT   = 'Your invitation · Singh Foundation Diwali 2026';
const FROM_NAME = 'Singh Foundation';

// Shown only in email apps that refuse HTML entirely (very rare).
const PLAIN_TEXT =
  'Singh Foundation invites you to Diwali 2026.\n\n' +
  'Friday, 30 October 2026, 6:00 - 9:00 pm\n' +
  '60 Gorge Rd, Trevallyn TAS 7250\n\n' +
  'Open your invitation and RSVP:\n' +
  'https://tassietech.github.io/sig/diwali-2026.html\n\n' +
  'Please RSVP by Thursday 1st October 2026. This is a private function.';


/** Sends one copy to yourself so you can check it before the real send. */
function sendTestToMe() {
  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    '[TEST] ' + SUBJECT,
    PLAIN_TEXT,
    { htmlBody: getHtml(), name: FROM_NAME }
  );
}


/** Sends to every guest in the Sheet who has not been sent to yet. */
function sendInvitations() {
  const html  = getHtml();
  const sheet = SpreadsheetApp.getActiveSheet();
  const rows  = sheet.getDataRange().getValues();

  let sent = 0, skipped = 0;

  for (let i = 1; i < rows.length; i++) {          // i = 1 skips the header row
    const email  = String(rows[i][0] || '').trim();
    const status = String(rows[i][1] || '').trim();

    if (!email || email.indexOf('@') === -1) { skipped++; continue; }
    if (status === 'SENT') { skipped++; continue; }

    if (MailApp.getRemainingDailyQuota() < 5) {
      sheet.getRange(i + 1, 2).setValue('QUOTA REACHED — run again tomorrow');
      break;
    }

    GmailApp.sendEmail(email, SUBJECT, PLAIN_TEXT, {
      htmlBody: html,
      name: FROM_NAME
    });

    sheet.getRange(i + 1, 2).setValue('SENT');
    SpreadsheetApp.flush();
    sent++;
    Utilities.sleep(1500);        // gentle pacing keeps you out of spam filters
  }

  SpreadsheetApp.getUi().alert(
    'Sent: ' + sent + '\nSkipped: ' + skipped +
    '\nRemaining quota today: ' + MailApp.getRemainingDailyQuota()
  );
}


/** Pulls the invitation markup out of the invitation.html file. */
function getHtml() {
  return HtmlService.createHtmlOutputFromFile('invitation').getContent();
}


/** Adds a menu to the Sheet so you can run this without opening the editor. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Diwali invites')
    .addItem('Send test to me', 'sendTestToMe')
    .addItem('Send to guest list', 'sendInvitations')
    .addToUi();
}

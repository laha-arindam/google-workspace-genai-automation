function autoReplyAndLogEmails() {
  /**
   * Searches for unread emails from the last 5 days with subjects starting with "ele 999" or "ELE 999",
   * sends an automatic reply, logs the email details in a Google Sheet, and marks them as read.
   *
   * Inputs: None (fetches emails automatically)
   * Outputs: Replies to emails, logs data in Google Sheets, and marks emails as read
   */
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dateLimit = new Date();
  dateLimit.setDate(dateLimit.getDate() - 5); // Get the date 5 days ago
  const threads = GmailApp.search('subject:("ele 999" OR "ELE 999") is:unread newer_than:5d');

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      if (!message.isUnread() || message.getDate() < dateLimit) return;
      sendReply(message);
      ensureHeaders(sheet);
      logEmail(sheet, message);
      message.markRead();
    });
  });
}

function sendReply(message) {
  /**
   * Sends an automatic response to the sender of the email.
   *
   * Inputs:
   * - message (GmailMessage): The email message object to which we are replying.
   *
   * Outputs:
   * - Sends an email reply to the sender.
   */
  const body = "I will get back to you.\n\nResit Sendag";
  GmailApp.sendEmail(message.getFrom(), "Re: " + message.getSubject(), body);
}

function ensureHeaders(sheet) {
  const headers = ["Timestamp", "From", "Subject", "Body Preview"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function logEmail(sheet, message) {
  /**
   * Logs the email details (date, sender, subject, and body) in the active Google Sheet.
   *
   * Inputs:
   * - sheet (Sheet): The Google Sheets object where data is logged.
   * - message (GmailMessage): The email message object containing details to be logged.
   *
   * Outputs:
   * - Appends a new row to the Google Sheet with the email's details.
   */
  sheet.appendRow([
    new Date(),
    message.getFrom(),
    message.getSubject(),
    message.getPlainBody().substring(0, 500) // Save as plain text, limit to 500 characters
  ]);
}


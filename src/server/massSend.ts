import { createDraftWithGmailAPI, getGmailTemplateFromDrafts } from "./helpers/gmail";


export interface MassSendConfig {
  recipientCol: string,
  emailSentCol: "Email Sent Status",
  confirmationEmailSubjectLine: string,
  classChosenCol: string,
  classMapping: Object
};

/**
 * Fetches the template from the given draft, fills in the template with appropriate column values, and sends the email (or creates a draft).
 * @param options The configuration options.
 * @param isDraft Whether to create a draft or send an email.
 * @param sheet The sheet to look for values from. Uses the active sheet by default.
 */
export function sendEmail(options: MassSendConfig, isDraft=true, sheet=SpreadsheetApp.getActiveSheet()) {
  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(options["emailSentCol"]);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails.
  // This array gets written out to the sheet only once the email has been sent.
  const out = [];

  // Array to record recipients that we are sending emails to.
  const recipients = [];

  const emailTemplate = getGmailTemplateFromDrafts(options["templateEmailSubjectLine"]);

  // Record each recipient in the list of recipients, and record the respective 'out' value.
  obj.forEach((row, rowIdx) => {
    if (row[options["emailSentCol"]] == '') {
      // Add this recipient to the list of recipients we want to send our email to.
      recipients.push(row[options["recipientCol"]]);
      // Email was sent, so cell should record the email sent date.
      out.push([new Date().toDateString()]);
    } else {
      // Email sent is already present, so cell should have whatever was there originally.
      out.push([row[options["emailSentCol"]]]);
    }
  });

  if (isDraft) {
    createDraftWithGmailAPI(recipients, options["confirmationEmailSubjectLine"], emailTemplate.message.text,emailTemplate.message.html, emailTemplate.attachments);
  } else {
    const recipientString = recipients.join(","); // Join the recipients into a comma-separated string of recipients.
    MailApp.sendEmail(recipientString, options["confirmationEmailSubjectLine"], emailTemplate.message.text, {
      htmlBody: emailTemplate.message.html,
      // bcc: 'a.bcc@email.com',
      // cc: 'a.cc@email.com',
      // from: 'an.alias@email.com',
      // name: 'name of the sender',
      // replyTo: 'a.reply@email.com',
      // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
      attachments: emailTemplate.attachments,
      inlineImages: emailTemplate.inlineImages
    });
  }
  
  // Update the 'Email Sent' column with the 'out' values to mark the emails as sent.
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
}

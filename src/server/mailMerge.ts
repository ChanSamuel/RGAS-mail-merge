import { createDraftWithGmailAPI, getGmailTemplateFromDrafts } from "./helpers/gmail";

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 */
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return  JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}


export interface MailMergeConfig {
  recipientCol: string,
  emailSentCol: "Email Sent Status",
  confirmationEmailSubjectLine: string,
  templateEmailSubjectLine: string
};

/**
 * Fetches the template from the given draft, fills in the template with appropriate column values, and sends the email (or creates a draft).
 * @param options The configuration options.
 * @param isDraft Whether to create a draft or send an email.
 * @param sheet The sheet to look for values from. Uses the active sheet by default.
 */
export function sendEmails(options: MailMergeConfig, isDraft=true, sheet=SpreadsheetApp.getActiveSheet()) {
  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  const emailSentColIdx = heads.indexOf(options["emailSentCol"]);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx) {
    // Only sends emails if email sent cell is blank and not hidden by a filter
    if (row[options["emailSentCol"]] == '') {
      const emailTemplate = getGmailTemplateFromDrafts(options["templateEmailSubjectLine"]);

      const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

      if (isDraft) {
        createDraftWithGmailAPI(row[options["recipientCol"]], options["confirmationEmailSubjectLine"], msgObj.text, msgObj.html, emailTemplate.attachments);
        
        // Draft is created, so write back whatever was there originally.
        out.push([row[options["emailSentCol"]]]);
      } else {
        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        MailApp.sendEmail(row[options["recipientCol"]], options["confirmationEmailSubjectLine"], msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bcc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Email was sent, so edit cell to record email sent date.
        out.push([new Date().toDateString()]);
      }
    } else {
      // Email sent is already present, so write back whatever was there originally.
      out.push([row[options["emailSentCol"]]]);
    }
  });
  
  // Update the 'Email Sent' column.
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
}

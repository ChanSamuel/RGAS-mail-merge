
/**
 * Creates a draft email using the Gmail API and MIMEText library.
 * This is neccessary since GmailApp does not support creation of drafts with new unicode characters (emojis, etc).
 */
function createDraftWithGmailAPI(recipientAddr, subject, textMsg, htmlMsg, senderName=Session.getActiveUser().getEmail(), senderAddr=Session.getActiveUser().getEmail(), attachments=[]) {
    const { message } = MimeText;
    message.setSender({
      name: senderName,
      addr: senderAddr,
    });
  
    const me = Session.getActiveUser().getEmail();
  
    message.setRecipient(recipientAddr);
    message.setSubject(subject);
    message.setMessage(textMsg, 'text/plain');
    message.setMessage(htmlMsg, 'text/html');
    message.setAttachments(attachments);
  
    const raw = message.asEncoded();
    Gmail.Users.Drafts.create({ message: { raw: raw } }, me);
  }
  
  function getColumnValuesByName(colName, ignoreEmptyValues=true, unique=true) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const data = sheet.getRange("A1:1").getValues();
    const col = data[0].indexOf(colName);
    if (col != -1) {
      var vals = sheet.getRange(2,col+1,sheet.getMaxRows()).getValues();
      vals = vals.map((e) => e[0]); // Convert from 2D array of 1 elements, to a 1D array of 'n' elements.
      if (ignoreEmptyValues) {
        vals = vals.filter((e) => e !== ""); // Remove empty strings.
      }
      if (unique) { // https://stackoverflow.com/questions/2218999/how-to-remove-all-duplicates-from-an-array-of-objects
        vals = [...new Set(vals)];
      }
      return vals;
    }
    throw Error(`Column ${colName} not found in active spreadsheet`);
  }
  
  function getColumnNames(ignoreEmptyHeaders=true) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getDisplayValues();
    const heads = data.shift(); // Assumes row 1 contains our column headings
    if (ignoreEmptyHeaders) {
      return heads.filter((e) => e != null);
    }
    return heads;
  }
  
  
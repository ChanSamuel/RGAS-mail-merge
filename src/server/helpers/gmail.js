
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
  
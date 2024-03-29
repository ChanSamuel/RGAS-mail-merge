import { createMimeMessage } from "mimetext"

/**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subjectLine to search for draft message
   * @return {object} GmailDraft object
  */
function subjectFilter_(subjectLine){
  return function(element) {
    if (element.getMessage().getSubject() === subjectLine) {
      return element;
    }
  }
}

/**
 * Creates a draft email using the Gmail API and MIMEText library.
 * This is neccessary since GmailApp does not support creation of drafts with new unicode characters (emojis, etc).
 * 
 * @param recipientAddr The email address of the recipient.
 * @param subject The subject line of the email.
 * @param textMsg The message to send in pure text format.
 * @param htmlMsg The message to send in html format.
 * @param senderName The name of the sender (defaults to GAS active user's email).
 * @param senderAddr The address of the sender (defaults to GAS active user's email)
 * @param attachments A list of attachment objects to add.
 */
export function createDraftWithGmailAPI(recipientAddr: string, subject: string, textMsg: string, htmlMsg: string, attachments: GoogleAppsScript.Gmail.GmailAttachment[], senderName=Session.getActiveUser().getEmail(), senderAddr=Session.getActiveUser().getEmail()) {
    const message = createMimeMessage();
    message.setSender({
      name: senderName,
      addr: senderAddr,
    });
  
    const me = Session.getActiveUser().getEmail();
  
    message.setRecipient(recipientAddr);
    message.setSubject(subject);
    message.addMessage({data: textMsg, contentType: 'text/plain'});
    message.addMessage({data: htmlMsg, contentType: 'text/html'});

    // TODO: Implement attachments.
    /*
    for (let i = 0; i < attachments.length; i++) {
      const atchmnt = attachments[i];
      message.addAttachment({
        filename: atchmnt.filename,
        contentType: atchmnt.contentType,
        data: atchmnt.data
      });
    }
    */
  
    const raw = message.asEncoded();
    Gmail.Users.Drafts.create({ message: { raw: raw } }, me);
}


/**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subjectLine to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
export function getGmailTemplateFromDrafts(subjectLine){
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subjectLine))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody(); 

    // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subjectLine, text: msg.getPlainBody(), html:htmlBody}, 
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }
}
  
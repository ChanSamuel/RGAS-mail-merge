import { sendEmails } from "./mailMerge"
import { sendMassEmail } from "./massSend";

export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('📜 Send emails', 'openSendEmailDialog')
    .addItem("Test Mail Merge", "testMailMerge")
    .addItem("Test Mass Send", "testMassSend");

  menu.addToUi();
};

export const openSendEmailDialog = () => {
  // Create the dialog.
  const html = HtmlService.createHtmlOutputFromFile('dialog-send-emails')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send emails');
};

export const testMailMerge = () => {
  const opts = {
    recipientCol: "Email:",
    emailSentCol: "Email Sent Status",
    confirmationEmailSubjectLine: "TESTING: Dance class conf email",
    templateEmailSubjectLine: "FOK Dance Class Template 1"
  }
  sendEmails(opts, true);
  SpreadsheetApp.getUi().alert("Email sent");
};

export const testMassSend = () => {
  const opts = {
    recipientCol: "Email:",
    emailSentCol: "Email Sent Status",
    confirmationEmailSubjectLine: "TESTING: Dance class conf email",
    templateEmailSubjectLine: "FOK Dance Class Template 1",
    classChosenCol: "Which class(es) will you be attending:",
    classMapping: {}
  }
  sendMassEmail(opts, true);
  SpreadsheetApp.getUi().alert("Email sent");
};


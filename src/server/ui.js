
export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('📜 Send emails', 'openSendEmailDialog');

  menu.addToUi();
};


export const openSendEmailDialog = () => {
  // Create the dialog.
  const html = HtmlService.createHtmlOutputFromFile('dialog-send-emails')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send emails');
};


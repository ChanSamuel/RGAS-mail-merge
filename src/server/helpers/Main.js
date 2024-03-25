// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
Developer email: mastersamuelchan@gmail.com
TODO: Change code to support 'n' different classes
TODO: On screen UI Form for configuring Mail Merge.
TODO: Send/create draft for multiple recipients.
TODO: Convert to React App.
TODO: Fix long wait times for UI config.
TODO: Support inline images and attachements and other extra email stuff.
TODO: Fix bug where date shows as month/day/year, rather than full date with weekday.
*/

/**
 * @OnlyCurrentDoc
*/

/**
 * Change these to match the column names you are using for email 
 * recipient addresses column, email sent column, and other options.
*/
var options = { // Declaring default options
    recipientCol: undefined,
    emailSentCol: "Email Sent Status",
    classChosenCol: undefined,
    classMapping: {},
    templateSubjectLineClass1: undefined,
    templateSubjectLineClass2: undefined,
    confirmationEmailSubjectLine: undefined,
  };
  
  /** 
   * Creates the menu items as a dropdown menu in Google Sheets.
   */
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Mail Merge')
        .addItem('📜 Create drafts', 'createDraftEmailsUIOption')
        .addToUi();
  }
  
  /**
   * UI Option for triggering a creation of drafts from sheet data.
   */
  function createDraftEmailsUIOption() {
    // Preconditions
    createEmailSentColsIfNotExists();
    setOptionsFromConfigSheet();

    collectTemplateAndExecuteAction("createDraft");
  }
  
  /**
   * UI Option for triggering a sending of emails from sheet data.
  */
  function sendEmailsUIOption() {
    // Preconditions.
    createEmailSentColsIfNotExists();
    setOptionsFromConfigSheet();
  
    // Confirm user actually wants to send emails.
    const ui = SpreadsheetApp.getUi();
    const confirm = ui.alert(
       'Please confirm',
       `You are about to send emails to all recipients who have not already been sent an email.\n
        Are you sure you want to continue?`,
        ui.ButtonSet.YES_NO);
    if (confirm === ui.Button.NO) {return;} // Operation was cancelled, so exit immediately.
  
    collectTemplateAndExecuteAction("sendEmail");
  }
  
  /**
   * Headless draft email triggering function.
   */
  function createDraftEmailsHeadless() {
    // Preconditions.
    createEmailSentColsIfNotExists();
    setOptionsFromConfigSheet();
  
    collectTemplateAndExecuteAction("createDraft");
  }
  
  /**
   * Headless send email triggering function.
   */
  function sendEmailsHeadless() {
    // Preconditions.
    createEmailSentColsIfNotExists();
    setOptionsFromConfigSheet();
  
    collectTemplateAndExecuteAction("sendEmail");
  }

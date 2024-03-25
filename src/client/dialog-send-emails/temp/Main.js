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
        .addItem("⚙️ Settings", "configureMMUIOption")
        .addToUi();
  }
  
  /**
   * UI Option for triggering a creation of drafts from sheet data.
   */
  function createDraftEmailsUIOption() {
    // Preconditions
    createEmailSentColsIfNotExists();
    setOptionsFromConfigSheet();
  
    // Confirm user actually wants to draft emails.
    const ui = SpreadsheetApp.getUi();
    var confirm = ui.alert(
       "Please confirm",
       `You are about to create draft emails for all recipients who have not already been sent an email.\n
        Are you sure you want to continue?`,
        ui.ButtonSet.YES_NO);
    if (confirm === ui.Button.NO) {return;} // Operation was cancelled, so exit immediately.
  
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
  
  function configureMMUIOption() {
    const htmlTemplate = HtmlService.createTemplateFromFile('settingsFrontPage.html');
    const htmlOutput = htmlTemplate.evaluate().setHeight(1000).setWidth(1000);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Settings');
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
  
  /**
   * Sends emails from sheet data.
   * @param {Sheet} sheet to read data from
  */
  function collectTemplateAndExecuteAction(action="sendEmail", sheet=SpreadsheetApp.getActiveSheet()) {
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
  
    // Creates an array to record sent emails
    const out = [];
  
    // Loops through all the rows of data
    obj.forEach(function(row, rowIdx){
      // Only sends emails if email sent cell is blank and not hidden by a filter
      if (row[options["emailSentCol"]] == '') {
        // Find the correct Email Template to send based on the class chosen.
        const classChosenRaw = row[options["classChosenCol"]];
  
        const classChosen = options["classMapping"][classChosenRaw];
        // Gets the draft Gmail message to use as a template
        var emailTemplate = undefined;
        if (classChosen === "Class 1") {
          emailTemplate = getGmailTemplateFromDrafts_(options["templateSubjectLineClass1"]);
        } else if (classChosen === "Class 2") {
          emailTemplate = getGmailTemplateFromDrafts_(options["templateSubjectLineClass2"]);
        } else {
          throw new Error(`No class mapping found for ${classChosenRaw}`);
        }
  
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
  
        if (action === "createDraft") {
          createDraftWithGmailAPI(row[options["recipientCol"]], options["confirmationEmailSubjectLine"], msgObj.text, msgObj.html, attachments=emailTemplate.attachments);
          
          // Draft is created, so write back whatever was there originally.
          out.push([row[options["emailSentCol"]]]);
        } else if (action === "sendEmail"){
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
        } else {
          throw new Error(`Action '${action}' not recognized`);
        }
      } else {
        // Email sent is already present, so write back whatever was there originally.
        out.push([row[options["emailSentCol"]]]);
      }
    });
    
    // Update the 'Email Sent' column.
    sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  }
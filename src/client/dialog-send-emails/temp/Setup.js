/**
 * Set the options given by the 'MM Config' sheet.
 */
function setOptionsFromConfigSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MM Config");
    // Assume columns will always be in that order.
    const class1 = sheet.getRange("A2").getValue();
    const class2 = sheet.getRange("B2").getValue();
    const recipientCol = sheet.getRange("C2").getValue();
    const classChosenCol = sheet.getRange("D2").getValue();
    const templateSubjectLineClass1 = sheet.getRange("E2").getValue();
    const templateSubjectLineClass2 = sheet.getRange("F2").getValue();
    const confirmationEmailSubjectLine = sheet.getRange("G2").getValue();
    options["classMapping"][class1] = "Class 1";
    options["classMapping"][class2] = "Class 2";
    options["recipientCol"] = recipientCol;
    options["classChosenCol"] = classChosenCol;
    options["templateSubjectLineClass1"] = templateSubjectLineClass1;
    options["templateSubjectLineClass2"] = templateSubjectLineClass2;
    options["confirmationEmailSubjectLine"] = confirmationEmailSubjectLine;
  }
  
  /**
   * Check if the Emails Sent column exists, if not then create it.
   */
  function createEmailSentColsIfNotExists() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];
    const headerRow = sheet.getRange("1:1");
    // Check if the Email Sent column exists. If so, return, otherwise, set it.
    if (headerRow.getDisplayValues()[0].includes(options["emailSentCol"])) {
      return;
    }
  
    // Set the cell in the next empty column with the Email Sent column name.
    const headerValues = headerRow.getDisplayValues()[0];
    // Loop through and find the next empty column. Assume it to be the last column plus one, by default.
    const lastCol = headerValues.length;
    var nextEmptyCol = lastCol + 1;
    for (let i = 0; i < headerValues.length; i++) {
      if (headerValues[i] === "") {
        nextEmptyCol = i + 1;
        break;
      } 
    }
    const cell = sheet.getRange(1, nextEmptyCol);
    cell.setValue(options["emailSentCol"]);
  }
  
  
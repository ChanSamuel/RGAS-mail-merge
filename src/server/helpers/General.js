
/**
 * Check if the Email Sent column exists, if not then create it.
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


const getSheets = () => SpreadsheetApp.getActive().getSheets();

const getActiveSheetName = () => SpreadsheetApp.getActive().getSheetName();

export const getSheetsData = () => {
  const activeSheetName = getActiveSheetName();
  return getSheets().map((sheet, index) => {
    const name = sheet.getName();
    return {
      name,
      index,
      isActive: name === activeSheetName,
    };
  });
};

export const addSheet = (sheetTitle) => {
  SpreadsheetApp.getActive().insertSheet(sheetTitle);
  return getSheetsData();
};

export const deleteSheet = (sheetIndex) => {
  const sheets = getSheets();
  SpreadsheetApp.getActive().deleteSheet(sheets[sheetIndex]);
  return getSheetsData();
};

export const setActiveSheet = (sheetName) => {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
  return getSheetsData();
};

export const getColumnValuesByName = (colName, ignoreEmptyValues=true, unique=true)  => {
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
  
export const getColumnNames = (ignoreEmptyHeaders=true) => {
    const sheet = SpreadsheetApp.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getDisplayValues();
    const heads = data.shift(); // Assumes row 1 contains our column headings
    if (ignoreEmptyHeaders) {
        return heads.filter((e) => e != null);
    }
    return heads;
}
  
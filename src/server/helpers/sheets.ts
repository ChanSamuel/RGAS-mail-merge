const getSheets_ = () => SpreadsheetApp.getActive().getSheets();

const getActiveSheetName_ = () => SpreadsheetApp.getActive().getSheetName();


export const getSheetsData = () => {
  const activeSheetName = getActiveSheetName_();
  return getSheets_().map((sheet, index) => {
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
  const sheets = getSheets_();
  SpreadsheetApp.getActive().deleteSheet(sheets[sheetIndex]);
  return getSheetsData();
};

export const setActiveSheet = (sheetName) => {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
  return getSheetsData();
};

export const getColumnValuesByName = (colName: string, ignoreEmptyValues=true, unique=true)  => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const data = sheet.getRange("A1:1").getValues();
    const col = data[0].indexOf(colName);
    if (col != -1) {
      const vals = sheet.getRange(2,col+1,sheet.getMaxRows()).getValues();
      let valsFlat = vals.flat();
      if (ignoreEmptyValues) {
        valsFlat = valsFlat.filter((e) => e !== ""); // Remove empty strings.
      }
      if (unique) {
        valsFlat = [...new Set(valsFlat)]; // Remove all duplicate values.
      }
      return valsFlat;
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
  
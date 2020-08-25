const DEFAULT_KEYS_COLUMN_ROW = 1;

function doGet(e: any): GoogleAppsScript.Content.TextOutput {
  const dataKeysColumnNumber: number = e.parameter.keys_column_row || 1;
  const dataStartRowNumber: number = e.parameter.start_row || 2;

  // e.parameterでURL QueryのObejctが取得できる
  const targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const resultObject = loadSpreadsheetToObjects(targetSpreadSheet, dataKeysColumnNumber, dataStartRowNumber);
  const jsonOut = ContentService.createTextOutput();
  //Mime TypeをJSONに設定
  jsonOut.setMimeType(ContentService.MimeType.JSON);
  //JSONテキストをセットする
  jsonOut.setContent(JSON.stringify(resultObject));
  return jsonOut;
}

function doPost(e: any): GoogleAppsScript.Content.TextOutput {
  const dataKeysColumnRow: number = e.parameter.keys_column_row || 1;
  const dataStartRowNumber: number = e.parameter.start_row || 2;
  const primaryKeyName = e.parameter.primary_key;

  const data = JSON.parse(e.postData.getDataAsString());
  const sheetNames = Object.keys(data);

  // e.parameterでURL QueryのObejctが取得できる
  const targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = targetSpreadSheet.getSheets();
  const newSheetNames = sheetNames.filter((sheetName) => sheets.every((sheet) => sheetName != sheet.getName()));
  for (const sheetName of newSheetNames) {
    const newSheet = targetSpreadSheet.insertSheet(sheetName);
    sheets.push(newSheet);
  }
  // Sheet名のdataを取り出す
  for (const sheetName of sheetNames) {
    const sheet = sheets.find((sheet) => sheet.getSheetName() === sheetName);
    if (!sheet) {
      continue;
    }
    // Sheet内のJSONData
    const sheetData = data[sheetName];
    const updateTargetRowsValuesList: { [n: number]: any }[] = [];
    if (sheet.getLastColumn() <= 0){
      const dataHeaderPairs: { [s: string]: any } = {}
      for (let i = 0; i < sheetData.length; ++i) {
        const rowData = sheetData[i];
        const rowKeys = Object.keys(rowData);
        for (let j = 0; j < rowKeys.length; ++j){
          const rowKey = rowKeys[j];
          if(!dataHeaderPairs[rowKey]){
            dataHeaderPairs[rowKey] = j + 1;
          }
        }
      }
      updateHeaderValues(sheet, dataHeaderPairs, dataKeysColumnRow);
    }
    const headerPairs = getKeyNumberPairs(sheet, dataKeysColumnRow);
    const headerValues = Object.values(headerPairs);
    let nextKeyNumber = headerValues.length > 0 ? Math.max(...headerValues) : 0;
    let maxColumnNumber = 1;
    // 1行分のObject
    for (let i = 0; i < sheetData.length; ++i) {
      const rowData = sheetData[i];
      const rowKeys = Object.keys(rowData);
      const updateTargetRowsValues: { [n: number]: any } = {};
      for (const rowKey of rowKeys) {
        // headerにないものKeyがきたらHeaderに追加する
        if (!headerPairs[rowKey]) {
          nextKeyNumber = nextKeyNumber + 1;
          headerPairs[rowKey] = nextKeyNumber;
        }
        // データの更新
        const headerColumnNumber = headerPairs[rowKey];
        updateTargetRowsValues[headerColumnNumber] = rowData[rowKey];
        if (maxColumnNumber < headerColumnNumber) {
          maxColumnNumber = headerColumnNumber;
        }
      }
      updateTargetRowsValuesList.push(updateTargetRowsValues);
    }
    // 変更すべきデータの行数の情報を取得
    const targetRowsRange = sheet.getRange(dataStartRowNumber, 1, sheetData.length, maxColumnNumber);
    const targetRowsValues = targetRowsRange.getValues();
    for (let i = 0; i < updateTargetRowsValuesList.length; ++i) {
      const updateColumnNumbers = Object.keys(updateTargetRowsValuesList[i]);
      let rowNumber;
      if (primaryKeyName) {
        rowNumber = targetRowsValues.findIndex(
          (rowValues) => rowValues[headerPairs[primaryKeyName] - 1] == updateTargetRowsValuesList[i][headerPairs[primaryKeyName]],
        );
      }
      if (!rowNumber || rowNumber < 0) {
        rowNumber = i;
      }
      for (const columnNumber of updateColumnNumbers) {
        targetRowsValues[rowNumber][columnNumber - 1] = updateTargetRowsValuesList[i][columnNumber];
      }
    }
    targetRowsRange.setValues(targetRowsValues);
  }
  const jsonOut = ContentService.createTextOutput();
  //Mime TypeをJSONに設定
  jsonOut.setMimeType(ContentService.MimeType.JSON);
  //JSONテキストをセットする
  jsonOut.setContent(JSON.stringify(data));
  return jsonOut;
}

function loadSpreadsheetToObjects(
  targetSpreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  dataKeysColumnNumber: number = 1,
  dataStartRowNumber: number = 2,
): { [s: string]: any } {
  const resultObject: { [s: string]: any } = {};
  for (const sheet of targetSpreadSheet.getSheets()) {
    const resultJsonObjects = [];
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    for (let row = dataStartRowNumber - 1; row < data.length; ++row) {
      const sheetData: { [s: string]: any } = {};
      const keys = data[0];
      for (let column = dataKeysColumnNumber - 1; column < keys.length; ++column) {
        sheetData[keys[column]] = data[row][column];
      }
      resultJsonObjects.push(sheetData);
    }
    resultObject[sheet.getSheetName()] = resultJsonObjects;
  }
  return resultObject;
}

function getKeyNumberPairs(
  targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
  headerKeysColumnRow: number = DEFAULT_KEYS_COLUMN_ROW,
): { [s: string]: number } {
  const keyNumberPairs: { [s: string]: number } = {};
  const headerRange = targetSheet.getRange(headerKeysColumnRow, 1, 1, targetSheet.getLastColumn());
  const headerValues = headerRange.getValues();
  if (headerValues[0]) {
    for (let i = 0; i < headerValues[0].length; ++i) {
      keyNumberPairs[headerValues[0][i]] = i + 1;
    }
  }
  return keyNumberPairs;
}

function updateHeaderValues(
  targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyNumberPairs: { [s: string]: number },
  headerKeysColumnRow: number = DEFAULT_KEYS_COLUMN_ROW,
): void {
  const keyArray = Object.keys(keyNumberPairs);
  keyArray.sort((a, b) => {
    if (keyNumberPairs[a] > keyNumberPairs[b]) {
      return 1;
    } else if (keyNumberPairs[a] < keyNumberPairs[b]) {
      return -1;
    } else {
      return 0;
    }
  });
  const columnNumbers: number[] = Object.values(keyNumberPairs);
  const maxColumnNumber = columnNumbers.length > 0 ? Math.max(...columnNumbers) : 1;
  const headerRange = targetSheet.getRange(headerKeysColumnRow, 1, 1, maxColumnNumber);
  headerRange.setValues([keyArray]);
}

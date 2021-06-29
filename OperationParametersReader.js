function getOperationParametersOnBatch(spreadsheetDBUrl, batchName) {
  let ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);

  let operationParametersSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.OPERATION_PARAMETERS);
  let tableObject = getTableFromSheet(operationParametersSheet, "operationParameters");
  let data = tableObject["data"];
  var col2Idx = tableObject["col2Idx"];

  var operationParameters = data.filter((row) => {
    return row[col2Idx["BatchName"]] == batchName || row[col2Idx["BatchName"]] == CONST.BATCH_NAME_ALL;
  });

  if (operationParameters.length > 0 ) {
    var operationParametersObj = {}
    for (let row of operationParameters) {
      operationParametersObj[row[col2Idx["Type"]]] = row[col2Idx["Value"]];
    }
    return operationParametersObj;
  } else {
    throw new BatchNotFoundException(`Batch with name "${batchName}" is not found!`);
  }
}

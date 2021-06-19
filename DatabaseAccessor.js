function getOperationParametersTable(spreadsheetDBUrl) {
    let ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);
  
    let operationParametersSheet = ssSource.getSheetByName("OperationParameters");
    let operationParametersTable = operationParametersSheet.getDataRange().getValues();
  
    return operationParametersTable;
  }
  
  function getOperationParametersOnBatch(spreadsheetDBUrl, batchName) {
    var operationParametersTable = getOperationParametersTable(spreadsheetDBUrl);
  
    var idx2Col = operationParametersTable[0];
    var col2Idx = {}
    for(let i = 0; i < idx2Col.length; i++)
      col2Idx[idx2Col[i]] = i
    Logger.log("[INFO]: col2idx for operationParameters is:\n%s", col2Idx)
  
    var operationParameters = operationParametersTable.filter((row) => {
      return row[col2Idx["BatchName"]] == batchName;
    });
  
    var operationParametersObj = {}
    for (let row of operationParameters) {
      operationParametersObj[row[col2Idx["Type"]]] = row[col2Idx["Value"]];
    }
    return operationParametersObj;
  }
  
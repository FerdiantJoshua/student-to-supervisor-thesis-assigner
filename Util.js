function getTableFromSheet(sheet, verboseTableName=null) {
  let table = sheet.getDataRange().getValues();
  let col2Idx = getCol2Idx(table[0]);
  if (verboseTableName != null)
    Logger.log("[INFO]: col2idx for table '%s' is:\n%s", verboseTableName, col2Idx)
  return {"data": table.slice(1), "col2Idx": col2Idx};
}

function isSheetNonData(sheetName) {
  let loweredSheetName = sheetName.toLowerCase();
  return loweredSheetName == CONST.SHEET_NAMES.TEMPLATE.toLowerCase();
}

function getCol2Idx(idx2Col) {
  let col2Idx = {}
  for(let i = 0; i < idx2Col.length; i++)
    col2Idx[idx2Col[i]] = i
  return col2Idx;
}

function groupByColumn(array2d, columnIdx, removeColumn=false) {
  groupedData = {}
  for(let i = 0; i < array2d.length; i++) {
    let row = array2d[i];
    let columnValue = !removeColumn ? row[columnIdx] : row.splice(columnIdx, 1);
    if (groupedData[columnValue] == null) {
      groupedData[columnValue] = [row];
    } else {
      groupedData[columnValue].push(row);
    }
  }
  return groupedData;
}

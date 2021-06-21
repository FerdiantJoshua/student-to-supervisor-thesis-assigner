function getAllFirstCellOfMarkedTables(spreadsheetAssignmentUrl) {
    var ssSource = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  
    var templateSheet = ssSource.getSheetByName("Template");
    var markerData = templateSheet.getRange("A:A").getValues(); // Markers are always put in column A
  
    var markersJSON = {};
    for(let [idx, row] of Object.entries(markerData)) {
      if (row[0] != "") {
        markersJSON[row[0]] = new Cell("B", idx + 1); // Google Spreadsheet is 1-indexed
      }
    }
    return markersJSON;
  }
  
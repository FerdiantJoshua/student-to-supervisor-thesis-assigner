function _getAllMarkerLocations(templateSheet) {
  var markerData = templateSheet.getRange("A:A").getValues(); // Markers are always put in column A

  var markersJSON = {};
  for(let [idx, row] of Object.entries(markerData)) {
    if (row[0] != "") {
      markersJSON[row[0]] = parseInt(idx) + 1; // Google Spreadsheet is 1-indexed
    }
  }
  return markersJSON;
}

function getPropsOfChosenStudentsTables(templateSheet) {
  let allMarkerLocations = _getAllMarkerLocations(templateSheet);
  let rowPositions_Header = {};
  Object.keys(allMarkerLocations)
    .filter(key => key.startsWith(CONST.MARKERS.SUPERVISION_LEVEL_PREFIX))
    .forEach(key => rowPositions_Header[key] = allMarkerLocations[key]);
  // e.g.: {"SUPERVISION_LEVEL_1": 13, "SUPERVISION_LEVEL_2": 21}

  let props = {};
  Object.entries(rowPositions_Header).forEach(
    (keyValPair) => {
      let [level, rowPos_Header] = keyValPair;
      level = level.split(CONST.MARKERS.SUPERVISION_LEVEL_PREFIX)[1];
      let firstRowPos = rowPos_Header + 1;
      let lastRowPos =  templateSheet.getRange(firstRowPos, CONST.TEMPLATE_FIRST_DATA_COLUMN, 1)
                                     .getNextDataCell(SpreadsheetApp.Direction.DOWN)
                                     .getRow();
      props[level] = {
        "firstRow": firstRowPos, "size": lastRowPos - firstRowPos
      };  // don't need to +1 to the size as the last row is not entry-able
    }
  );
  return props  // e.g.: {"1": {"firstRow": 13, "size:": 6}, "2": {"firstRow": 21, "size:": 4}}
}

function getPropsOfStudentQueueTable(templateSheet) {
  var allMarkerLocations = _getAllMarkerLocations(templateSheet);
  var rowPos_Header = allMarkerLocations[CONST.MARKERS.STUDENT_QUEUE];
  var header = templateSheet.getRange(
    rowPos_Header, CONST.TEMPLATE_FIRST_DATA_COLUMN, 1, templateSheet.getLastColumn()
  ).getValues()[0].filter(
    (value) => {return value != ""}
  );
  var tableSize = templateSheet.getMaxRows() - rowPos_Header;
  return {
    "firstRow": rowPos_Header + 1,
    "header": header,
    "size": tableSize,
  };  // e.g.: {"firstRow": 24, "header": ["NRP - Nama", "KK 1", "KK 2", "KK 3", "Perkiraan Judul Tugas Akhir"], "size": 150}
}

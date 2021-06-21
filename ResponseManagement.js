const firstCell_StudentQueues = [
  new Cell("C", 25),
  // new Cell("B", 30),
  // new Cell("I", 30),
  // new Cell("P", 30)
]

const firstCell_ChosenStudentLists = [
  new Cell("C", 13),
  // new Cell("B", 14),
  // new Cell("B", 21)
]

function assignStudentsToSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  // TERMINATE IF the spreadsheets has NOT been PREPARED
  if (ssTarget.getNumSheets() <= 2) {
    throw new UnpreparedSpreadsheetException(
      `Spreadsheet has not been prepared yet! Please run operation "Prepare Response Sheets" first.`,
      {"spreadsheetAssignmentUrl": spreadsheetAssignmentUrl}
    );
  }

  // TERMINATE IF any of professor sheet's student queues IS NOT EMPTY
  var professorSheets = ssTarget.getSheets();
  for(let professorSheet of professorSheets) {
    if (_isSheetNonData(professorSheet.getName())) {
      continue;
    }
    for(var firstCell_StudentQueue of firstCell_StudentQueues) {
      let firstCellValue = professorSheet.getRange(firstCell_StudentQueue.getPos()).getValues();
      if (firstCellValue[0][0] != "") {
        throw new StudentsAlreadyAssignedException(
          `Students in spreadsheet are already assigned! If you WANT TO REASSIGN students, run operation "Clear All Professor Sheets" first.`,
          {"spreadsheetAssignmentUrl": spreadsheetAssignmentUrl}
        );
      }
    }
  }

  let formResponsesSheet = ssTarget.getSheetByName("Form Responses");
  let table = formResponsesSheet.getDataRange().getValues();

  var col2Idx = getCol2Idx(table[0]);
  Logger.log("[INFO]: col2idx for formResponses is:\n%s", col2Idx)

  // find every columns in "Form Responses Sheet" which starts with "pilihan dosen"
  var chosenProfessorColNames = Object.keys(col2Idx).filter(
    (key) => {return key.toLowerCase().startsWith("pilihan dosen")}
  );
  if (chosenProfessorColNames.length > firstCell_StudentQueues) {
    console.warn(
      `chosenProfessorColNames.length > firstCell_StudentQueues.length! (${chosenProfessorColNames.length} vs ${firstCell_StudentQueues.length})`
    );
  }
  console.log("chosenProfessorColNames is:", chosenProfessorColNames)
  
  let data = table.slice(1);
  for(var row of data) {
    var valuesToAppend = [
      [
        row[col2Idx["NRP - Nama"]],
        row[col2Idx["KK 1"]],
        row[col2Idx["KK 2"]],
        row[col2Idx["KK 3"]],
        row[col2Idx["Perkiraan Judul Tugas Akhir"]]
      ]
    ];
    // chosenProfessors = [row[col2Idx["Pilihan Dosen 1"]], row[col2Idx["Pilihan Dosen 2"]], row[col2Idx["Pilihan Dosen 3"]]]

    for(var i in chosenProfessorColNames) {
      chosenProfessor = row[col2Idx[chosenProfessorColNames[i]]];
      professorName = chosenProfessor.split(" - ")[0]
      professorSheet = ssTarget.getSheetByName(professorName)
      if (professorSheet == null) {
        console.warn(
          "[WARNING]: Sheet for professor name '%s' is not found (on student '%s'). Skipping append..", professorName, valuesToAppend[0][0]
        )
        continue;
      }
      let nextEmptyRowInColumn = firstCell_StudentQueues[i].getNextEmptyRow(professorSheet)

      let rangeToInsert = nextEmptyRowInColumn.getRangeNextNCols(valuesToAppend[0].length - 1)
      professorSheet.getRange(rangeToInsert).setValues(valuesToAppend)
      Logger.log("[INFO]: Successfully append '%s' to Sheet '%s!%s'", valuesToAppend, professorName, rangeToInsert)
    }
  }
}

function clearAllProfessorSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  var professorSheets = ssTarget.getSheets();
  if (professorSheets.length == 0)
    Logger.log("[INFO] No professor sheets to clear");

  for(var professorSheet of professorSheets) {
    Logger.log("[INFO] Clearing sheet: '%s'", professorSheet.getName())
    
    if (_isSheetNonData(professorSheet.getName())) {
      Logger.log("[INFO] Accessing non-data sheet: '%s'. Skipping..", professorSheet.getName())
      continue
    }

    for(var firstCell_StudentQueue of firstCell_StudentQueues) {
      let botRightCell_StudentQueue = firstCell_StudentQueue.getNextEmptyCol(professorSheet);
      botRightCell_StudentQueue.row = professorSheet.getMaxRows();
      let rangeToClear = firstCell_StudentQueue.getPos() + ":" + botRightCell_StudentQueue.getPos();
      professorSheet.getRange(rangeToClear).clear({contentsOnly: true});
    }
    for(var firstCell_ChosenStudentList of firstCell_ChosenStudentLists) {
      let rangeToClear = firstCell_ChosenStudentList.getPos() + ":" + firstCell_ChosenStudentList.getNextEmptyRow(professorSheet).getPos();
      professorSheet.getRange(rangeToClear).clear({contentsOnly: true});
    }
  }
}

function _isSheetNonData(sheetName) {
  let loweredSheetName = sheetName.toLowerCase();
  return loweredSheetName == "template" || loweredSheetName.includes("form responses");
}

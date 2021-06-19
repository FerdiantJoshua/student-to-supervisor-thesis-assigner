class Cell {
  constructor(col='A', row=1) {
    this.col = col.toUpperCase()
    this.row = row
  }

  getNumberRepr() {
    return [this.col.charCodeAt() - 64, this.row]
  }

  getPos() {
    return this.col + this.row
  }

  getNextNCols(n) {
    return new Cell(String.fromCharCode(this.col.charCodeAt() + n), this.row)
  }

  getNextNRows(n) {
    return new Cell(this.col, this.row + n)
  }

  getRangeNextNCols(n) {
    let endCell = this.getNextNCols(n)
    return this.getPos() + ":" + endCell.getPos()
  }

  getRangeNextNRows(n) {
    let endCell = this.getNextNRows(n)
    return this.getPos() + ":" + endCell.getPos()
  }

  getRangeFullCol() {
    return this.getPos() + ":" + this.col
  }

  getRangeFullRow() {
    return this.getPos() + ":" + this.row
  }

  getNextEmptyCol(sheet) {
    let fullRow = sheet.getRange(this.getRangeFullRow());
    let values = fullRow.getValues();
    let ct = 0;
    while ( values[0][ct] != "" ) {
      ct++;
    }
    return new Cell(String.fromCharCode(this.col.charCodeAt() + ct), this.row);
  }

  getNextEmptyRow(sheet) {
    let fullColumn = sheet.getRange(this.getRangeFullCol());
    let values = fullColumn.getValues();
    let ct = 0;
    while ( values[ct] && values[ct][0] != "" ) {
      ct++;
    }
    return new Cell(this.col, this.row + ct);
  }
}

const firstCell_StudentQueues = [
  new Cell("B", 24),
  // new Cell("B", 30),
  // new Cell("I", 30),
  // new Cell("P", 30)
]

const firstCell_ChosenStudentLists = [
  new Cell("B", 13),
  // new Cell("B", 14),
  // new Cell("B", 21)
]

function assignStudentsToSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  let formResponsesSheet = ssTarget.getSheetByName("Form Responses");
  let table = formResponsesSheet.getDataRange().getValues();

  var idx2Col = table[0];
  var col2Idx = {}
  for(var i = 0; i < idx2Col.length; i++)
    col2Idx[idx2Col[i]] = i
  Logger.log("[INFO]: col2idx for formResponses is:\n%s", col2Idx)

  // find every columns in "Form Responses Sheet" which starts with "pilihan dosen"
  var chosenProfessorColNames = Object.keys(col2Idx).filter(
    (key) => {return key.toLowerCase().startsWith("pilihan dosen")}
  );
  if (chosenProfessorColNames.length > firstCell_StudentQueues) {
    console.warn(`chosenProfessorColNames.length > firstCell_StudentQueues.length! (${chosenProfessorColNames.length} vs ${firstCell_StudentQueues.length})`);
  }
  console.log("chosenProfessorColNames is:", chosenProfessorColNames)
  
  let data = table.slice(1)
  for(var row of data) {
    var valuesToAppend = [
      [
        row[col2Idx["NRP - Nama"]],
        row[col2Idx["Pilihan Lingkup Penelitian 1"]],
        row[col2Idx["Pilihan Lingkup Penelitian 2"]],
        row[col2Idx["Pilihan Lingkup Penelitian 3"]],
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
    
    let loweredSheetName = professorSheet.getName().toLowerCase();
    if (loweredSheetName == "template" || loweredSheetName.includes("form responses")) {
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

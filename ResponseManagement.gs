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

  getPosNextNCols(n) {
    return new Cell(String.fromCharCode(this.col.charCodeAt() + n), this.row)
  }

  getPosNextNRows(n) {
    return new Cell(this.col, this.row + n)
  }

  getRangeNextNCols(n) {
    let endCell = this.getPosNextNCols(n)
    return this.getPos() + ":" + endCell.getPos()
  }

  getRangeNextNRows(n) {
    let endCell = this.getPosNextNRows(n)
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

const studentQueueCells = [
  new Cell("B", 30),
  new Cell("I", 30),
  new Cell("P", 30)
]

const chosenStudentListCells = [
  new Cell("B", 14),
  new Cell("B", 21)
]

function assignStudentsToSheets(spreadsheetTargetUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetTargetUrl);

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
  if (chosenProfessorColNames.length > studentQueueCells) {
    console.warn(`chosenProfessorColNames.length > studentQueueCells.length! (${chosenProfessorColNames.length} vs ${studentQueueCells.length})`);
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
      let nextEmptyCellInColumn = studentQueueCells[i].getNextEmptyRow(professorSheet)

      let rangeToInsert = nextEmptyCellInColumn.getRangeNextNCols(valuesToAppend[0].length - 1)
      professorSheet.getRange(rangeToInsert).setValues(valuesToAppend)
      Logger.log("[INFO]: Successfully append '%s' to Sheet '%s!%s'", valuesToAppend, professorName, rangeToInsert)
    }
  }
}

function clearAllProfessorSheets(spreadsheetTargetUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetTargetUrl);

  var professorSheets = ssTarget.getSheets().slice(2);
  if (professorSheets.length == 0) Logger.log("[INFO] No professor sheets to clear")
  for(var professorSheet of professorSheets) {
    Logger.log("[INFO] Clearing sheet: '%s'", professorSheet.getName())
    
    let loweredSheetName = professorSheet.getName().toLowerCase();
    if (loweredSheetName == "template" || loweredSheetName.includes("form responses")) {
      Logger.log("[INFO] Accessing non-data sheet: '%s'. Skipping..", professorSheet.getName())
      continue
    }

    for(var studentQueueCell of studentQueueCells) {
      let laststudentQueueCell = studentQueueCell.getNextEmptyCol(professorSheet);
      laststudentQueueCell.row = professorSheet.getMaxRows();
      let rangeToClear = studentQueueCell.getPos() + ":" + laststudentQueueCell.getPos();
      professorSheet.getRange(rangeToClear).clear({contentsOnly: true});
    }
    for(var chosenStudentListCell of chosenStudentListCells) {
      let rangeToClear = chosenStudentListCell.getPos() + ":" + chosenStudentListCell.getNextEmptyRow(professorSheet).getPos();
      professorSheet.getRange(rangeToClear).clear({contentsOnly: true});
    }
  }
}

function helloWorld(text) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Toast Message", text);
}

function assignStudentsToSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  // Fetch 'Student Queue Table' info from 'Template' sheet
  let templateSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);
  var props_StudentQueueTable = getPropsOfStudentQueueTable(templateSheet);
  console.log("[DEBUG] props_StudentQueueTable:", props_StudentQueueTable);

  // TERMINATE IF the spreadsheets has NOT been PREPARED
  if (ssTarget.getNumSheets() <= 2) {
    throw new UnpreparedSpreadsheetException(
      `Spreadsheet has not been prepared yet! Please run operation "Prepare Response Sheets" first.`,
      debug={"spreadsheetAssignmentUrl": spreadsheetAssignmentUrl}
    );
  }

  // TERMINATE IF any of professor sheet's student queues IS NOT EMPTY
  var professorSheets = ssTarget.getSheets();
  for(let professorSheet of professorSheets) {
    if (isSheetNonData(professorSheet.getName())) {
      continue;
    }
    let firstCellValue = professorSheet.getRange(props_StudentQueueTable.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN, 1).getValues();
    if (firstCellValue[0][0] != "") {
      throw new StudentsAlreadyAssignedException(
        `Students in spreadsheet are already assigned! If you WANT TO REASSIGN students, run operation "Clear All Professor Sheets" first.`,
        debug={"spreadsheetAssignmentUrl": spreadsheetAssignmentUrl}
      );
    }
  }

  // Read data from sheet 'Form Responses'
  let formResponsesSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.FORM_RESPONSES);
  let tableObject = getTableFromSheet(formResponsesSheet, "formResponses");
  let data = tableObject["data"];
  var col2Idx = tableObject["col2Idx"];

  let emailIdx = col2Idx["Email Address"];
  let nrpNameIdx = col2Idx["NRP - Nama"];
  let i = 0
  while (i < data.length) {
    let row = data[i];
    let nrpFromEmail = row[emailIdx].split('@', 1)[0];
    let nrpFromName = row[nrpNameIdx].split(' - ', 1)[0];
    if ( nrpFromEmail != nrpFromName) {
      console.warn(
        "[WARNING] Student with name '%s' is omitted due to NRP mismatch '%s'. ('%s' vs '%s')",
        row[nrpNameIdx], row[emailIdx], nrpFromName, nrpFromEmail
      )
      data.splice(i, 1);
    } else {
      i += 1;
    }
  }
  
  // Assign students to professors
  let responsesGroupedByProfessor = groupByColumn(data, col2Idx[CONST.COL_NAMES.PILIHAN_DOSEN]);
  for(let [professorName, studentResponseList] of Object.entries(responsesGroupedByProfessor)) {
    professorName = professorName.split(" - ")[0]
    professorSheet = ssTarget.getSheetByName(professorName)
    if (professorSheet == null) {
      console.warn(
        "[WARNING]: Sheet for professor name '%s' is not found (chosen by students '%s'). Skipping append..", professorName, studentResponseList
      )
      continue;
    }

    let nDataToInsert = Math.min(studentResponseList.length, props_StudentQueueTable.size);
    console.log("[DEBUG] nDataToInsert:", nDataToInsert);

    let rangeToInsert = professorSheet.getRange(
      props_StudentQueueTable.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN,
      nDataToInsert, props_StudentQueueTable.header.length
    )
    
    let valuesToInsert = []
    for(let i in studentResponseList) {
      if (i >= nDataToInsert) {
        break;
      }
      studentResponse = studentResponseList[i];
      valuesToInsert.push(
        props_StudentQueueTable.header.map((columnName) => {
          return studentResponse[col2Idx[columnName]]
        }
      ))
    }
    
    professorSheet.getRange(rangeToInsert.getA1Notation()).setValues(valuesToInsert)
    Logger.log("[INFO]: Successfully inserted '%s' to Sheet '%s!%s'", valuesToInsert, professorName, rangeToInsert)
  }
}

function clearAllProfessorSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  // Fetch 'Student Queue Table' info from 'Template' sheet
  let templateSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);
  var props_StudentQueueTable = getPropsOfStudentQueueTable(templateSheet);
  console.log("[DEBUG] props_StudentQueueTable:", props_StudentQueueTable);

  // Clear professor sheets
  var professorSheets = ssTarget.getSheets();
  if (professorSheets.length == 0) Logger.log("[INFO] No professor sheets to clear");
  for(var professorSheet of professorSheets) {
    professorSheetName = professorSheet.getName();
    if (isSheetNonData(professorSheetName)) {
      Logger.log("[INFO] Accessing non-data sheet: '%s'. Skipping..", professorSheetName);
      continue
    }
    Logger.log("[INFO] Clearing sheet: '%s'", professorSheetName);

    let rangeToClear = professorSheet.getRange(
      props_StudentQueueTable.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN,
      props_StudentQueueTable.size, props_StudentQueueTable.header.length
    )
    professorSheet.getRange(rangeToClear.getA1Notation()).clearContent();
  }
}

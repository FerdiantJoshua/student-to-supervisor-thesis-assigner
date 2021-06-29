function assignStudentsToSheets(spreadsheetDBUrl, spreadsheetAssignmentUrl, spreadsheetResponsesUrl, formResponsesSheetName) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  // Fetch 'Student Queue Table' info from 'Template' sheet
  let templateSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);
  var props_StudentQueueTable = getPropsOfStudentQueueTable(templateSheet);
  console.log("[DEBUG] props_StudentQueueTable:", props_StudentQueueTable);


  // TERMINATE IF the spreadsheets has NOT been PREPARED
  if (ssTarget.getNumSheets() <= 1) {
    throw new UnpreparedSpreadsheetException(
      `Spreadsheet has not been prepared yet! Please run operation "Generate Professor Sheets" first.`,
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


  // Fetch "Supervisions" table from sheet
  let ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);
  let supervisionSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.SUPERVISIONS);
  let supervisionsGroupedByStudent = getSupervisionsOnLevel(supervisionSheet, level="all", groupByStudent=true);

  // Read data from sheet 'Form Responses'
  let formResponsesSheet = SpreadsheetApp.openByUrl(spreadsheetResponsesUrl).getSheetByName(formResponsesSheetName);
  let tableObject = getTableFromSheet(formResponsesSheet, "formResponses");
  let {data, col2Idx} = tableObject;
  let {"NRP - Nama": nrpNameIdx, "Email Address": emailIdx, additionalColName: currSupervisorsIdx} = col2Idx;

  _omitStudentsWithInvalidEmail(data, emailIdx, nrpNameIdx);

  // The "Student Queue Table" has additional column which doesn't come from "Form Responses", but from "Supervision Table" in spreadsheetDB
  currSupervisorsIdx = _insertExistingSupervisorsToStudentResponses(data, nrpNameIdx, supervisionsGroupedByStudent);
  col2Idx[CONST.COL_NAMES.PEMBIMBING_LAIN] = currSupervisorsIdx;
  console.log('PURUYEAH', data)
  
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
        props_StudentQueueTable.header.map(columnName => studentResponse[col2Idx[columnName]])
      )
    }
    
    professorSheet.getRange(rangeToInsert.getA1Notation()).setValues(valuesToInsert)
    Logger.log("[INFO]: Successfully inserted '%s' to Sheet '%s!%s'", valuesToInsert, professorName, rangeToInsert)
  }
}

// Omit students whose email's username != his/her NRP
function _omitStudentsWithInvalidEmail(data, emailIdx, nrpNameIdx) {
  let i = 0
  while (i < data.length) {
    let row = data[i];
    let nrpFromEmail = row[emailIdx].split('@', 1)[0];
    let nrpFromName = row[nrpNameIdx].split(' - ', 1)[0];
    if ( nrpFromEmail != nrpFromName) {
      console.warn(
        "[WARNING] Student with name '%s' is omitted due to 'NRP & email' mismatch '%s'. ('%s' vs '%s')",
        row[nrpNameIdx], row[emailIdx], nrpFromName, nrpFromEmail
      )
      data.splice(i, 1);
    } else {
      i += 1;
    }
  }
}

// Insert students' current supervisor in-place; Return index of the new column
function _insertExistingSupervisorsToStudentResponses(data, nrpNameIdx, supervisionsGroupedByStudent) {
  console.log("supervisionsGroupedByStudent", supervisionsGroupedByStudent);
  for(let i = 0; i < data.length; i++) {
    let studentName = data[i][nrpNameIdx].split(' - ', 2)[1];
    let supervisors = supervisionsGroupedByStudent[studentName];
    data[i].push(
      supervisors != null ? supervisors.map(val => `${val[1]}: ${val[0]}`).join("\n") : ""
    )
  }
  return data[0].length - 1;
}

function clearAllStudentQueues(spreadsheetAssignmentUrl) {
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

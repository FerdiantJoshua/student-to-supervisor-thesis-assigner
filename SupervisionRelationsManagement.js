// "ssSource" and "ssTarget" VALUES ARE REVERSED in this function!
function saveStudentProfessorRelations(spreadsheetDBUrl, spreadsheetAssignmentUrl, forceSave=false) {
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  // Fetch 'Chosen Student List Table' info from 'Template' sheet
  let templateSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);
  let props_ChosenStudentsTables = getPropsOfChosenStudentsTables(templateSheet);
  console.log("[DEBUG] props_ChosenStudentsTables:", props_ChosenStudentsTables);
  let validStudentValuesCol = CONST.TEMPLATE_VALID_STUDENTS_COL_GETTER(templateSheet);
  console.log("[DEBUG] validStudentValuesCol:", validStudentValuesCol);

  // Fetch "Supervisions" table from sheet
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetDBUrl);
  var supervisionsSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.SUPERVISIONS);

  var professorSheets = ssSource.getSheets();
  var valuesToInsert = [];
  for(let professorSheet of professorSheets) {
    professorSheetName = professorSheet.getName();
    if (isSheetNonData(professorSheetName)) {
      Logger.log("[INFO] Found non-data sheet: '%s'. Skipping..", professorSheetName);
      continue
    }
    Logger.log("[INFO] Reading 'Chosen Students Table' from professorSheet: '%s'", professorSheetName);

    for(let level of Object.keys(props_ChosenStudentsTables)) {
      let props = props_ChosenStudentsTables[level];
      let rowValues = [];

      if (!forceSave) {
        let supervisionsGroupedByProfessor = getSupervisionsOnLevel(supervisionsSheet, level, groupByStudent=false);
        let prevNStudents = _getNStudentsOfProfessorOnLevel(supervisionsGroupedByProfessor, professorSheetName, level);

        let prevStudentRange = currStudentRange = validStudentNames = null;
        let [prevValues, currValues] = [[],[]];
        // To respect the 'data validation' in "Chosen Students Tables", values from prev batch are handled differently from curr batch
        if (prevNStudents != 0) {
          prevStudentRange = professorSheet.getRange(props.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN, prevNStudents);
          prevValues = prevStudentRange.getValues();
        }
        if (prevNStudents != props.size) {
          currStudentRange = professorSheet.getRange(props.firstRow + prevNStudents, CONST.TEMPLATE_FIRST_DATA_COLUMN, props.size - prevNStudents);
          currValues = currStudentRange.getValues();
          validStudentNames = professorSheet.getRange(1, validStudentValuesCol, CONST.FIRST_N_VALID_STUDENTS + 1)   // +1 for reserved "CLOSED"
                                            .getValues().flat();
          _validateStudentValues(currValues, validStudentNames);
          _sortPutClosedSlotsFirst(currValues);

          currStudentRange.setValues(currValues);
          currStudentRange.setBackground("cyan");
        }
        rowValues = prevValues.concat(currValues);
      } else {
        studentRange = professorSheet.getRange(props.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN, props.size);
        studentRange.setBackground("cyan");
        rowValues = studentRange.getValues();
      }

      let rowsToInsert = rowValues
                          .filter(studentName => studentName[0] != "")
                          .map(studentName => [_extractStudentName(studentName[0]), professorSheetName, parseInt(level)]);
      // console.log("[DEBUG] rowsToInsert", rowsToInsert);
      valuesToInsert.push(...rowsToInsert);
    }
  }
  console.log("[DEBUG] valuesToInsert:", valuesToInsert);

  // Clear and update "Supervisions" table
  supervisionsSheet.getRange(2, 1, supervisionsSheet.getMaxRows(), supervisionsSheet.getMaxColumns()).clearContent();
  if (valuesToInsert.length > 0) {
    supervisionsSheet.getRange(2, 1, valuesToInsert.length, valuesToInsert[0].length).setValues(valuesToInsert);
  } else {
    Logger.log("[INFO] All professors' 'Chosen Student List Table' is empty. No values were inserted.")
  }
}

function _getNStudentsOfProfessorOnLevel(supervisionsGroupedByProfessor, professorName, level) {
  let count = 0
  let students = supervisionsGroupedByProfessor[professorName];
  if (students != null) {
    students.forEach(student => count += student[1] == level ? 1 : 0)
  }
  return count;
}

function _validateStudentValues(studentValues, validStudentNames) {
  let validStudentNamesSet = new Set(validStudentNames);
  let duplicationSet = new Set();
  let i = 0;
  while (i < studentValues.length) {
    name = studentValues[i][0];
    // console.log("[DEBUG] name, validSet, dupSet =", name, validStudentNamesSet, duplicationSet);
    if (name != "" && name != CONST.CLOSED_SLOT_VALUE) {
      if (!validStudentNamesSet.has(name)) {
        studentValues[i][0] = CONST.CLOSED_SLOT_VALUE;
      } else if (duplicationSet.has(name)) {
        studentValues[i][0] = "";
      }
    }
    duplicationSet.add(name);
    i += 1;
  }
}

// function _turnOnStrictDataValidation(range) {
//   let rules = range.getDataValidations();
//   for (let i = 0; i < rules.length; i++) {
//     for (let j = 0; j < rules[i].length; j++) {
//       let criteria = rules[i][j].getCriteriaType();
//       let args = rules[i][j].getCriteriaValues();
//       rules[i][j] = rules[i][j].copy().withCriteria(criteria, args).setAllowInvalid(false).build();
//     }
//   }
//   range.setDataValidations(rules);
// }

function _sortPutClosedSlotsFirst(values_Array2d) {
  let closedSlots = [];
  let openSlots = [];
  let i = 0;
  while (i < values_Array2d.length) {
    if (values_Array2d[i][0] == CONST.CLOSED_SLOT_VALUE) {
      closedSlots.push(values_Array2d.splice(i, 1));
    } else if (values_Array2d[i][0] == "") {
      openSlots.push(values_Array2d.splice(i, 1));
    } else {
      i += 1;
    }
  }
  values_Array2d.unshift(...closedSlots);
  values_Array2d.push(...openSlots);
}

function _extractStudentName(studentNameWithNRP) {
  if (studentNameWithNRP != CONST.CLOSED_SLOT_VALUE) {
    return studentNameWithNRP.split(" - ", 2)[1]
  }
  return studentNameWithNRP
}

function getSupervisionsOnLevel(supervisionSheet, level="all", groupByStudent=true) {
  let tableObject = getTableFromSheet(supervisionSheet, CONST.SHEET_NAMES.SUPERVISIONS);
  let {data, col2Idx} = tableObject;
  let {"Student": studentIdx, "Professor": professorIdx, "Level": levelIdx} = col2Idx;

  let supervisions = level == "all" ? data : data.filter(row => row[levelIdx] == level);
  let groupColumnIdx = groupByStudent ? studentIdx : professorIdx;
  let supervisionsGrouped =  groupByColumn(supervisions, groupColumnIdx, removeColumn=true);
  return supervisionsGrouped;  // e.g. {"<student_name_1>": [[<professor_name_1>, <level>], [<professor_name_2>, <level>]], ...}
}

function getMaxStudentsBySupervisionLevel(supervisionLevelsSheet) {
  let tableObject = getTableFromSheet(supervisionLevelsSheet, CONST.SHEET_NAMES.SUPERVISION_LEVELS);
  let {data, col2Idx} = tableObject;
  let {"Level": levelIdx, "Max Students": maxStudentsIdx} = col2Idx;

  let result = {};
  data.forEach(row => result[row[levelIdx]] = parseInt(row[maxStudentsIdx]));
  return result;  // e.g. {"1": 6, "2": 4, "3": 2}
}

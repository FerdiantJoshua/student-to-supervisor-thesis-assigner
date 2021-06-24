// "ssSource" and "ssTarget" VALUES ARE REVERSED in this function!
function saveStudentProfessorRelations(spreadsheetDBUrl, spreadsheetAssignmentUrl) {
    var ssSource = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  
    // Fetch 'Chosen Student List Table' info from 'Template' sheet
    let templateSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);
    var props_ChosenStudentsTables = getPropsOfChosenStudentsTables(templateSheet);
    console.log("[DEBUG] props_ChosenStudentsTables:", props_ChosenStudentsTables);
  
    var chosenStudentsTablesA1Notations = {};
    Object.keys(props_ChosenStudentsTables).forEach(
      (level) => {
        let props = props_ChosenStudentsTables[level];
        chosenStudentsTablesA1Notations[level] = templateSheet
                                                  .getRange(props.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN, props.size)
                                                  .getA1Notation();
      }
    );  // e.g.: {"1":"C13:C18", "2":"C21:C26"}
  
    var professorSheets = ssSource.getSheets();
    if (professorSheets.length == 0)
      Logger.log("[INFO] No professor sheets to process");
  
    // var professorSheetNames = []
    // for(var professorSheet of professorSheets) {
    //   professorSheetName = professorSheet.getName();
    //   if (isSheetNonData(professorSheetName)) {
    //     // Logger.log("[INFO] Found non-data sheet: '%s'. Skipping..", professorSheetName);
    //     continue
    //   }
    //   // Logger.log("[INFO] Found professorSheet: '%s'", professorSheetName);
    //   professorSheetNames.push(professorSheetName);
    // }
    // console.log("[DEBUG] professorSheetNames", professorSheetNames);
  
    // var rangesToRead = professorSheetNames.flatMap(
    //   (professorSheetName) => {
    //     return chosenStudentsTablesA1Notations.map(a1Notation => `${professorSheetName}!${a1Notation}`)
    //     }
    // );  // // e.g.: ["Prof1!C13:C18", "Prof1!C21:C26", "Prof2!C13:C18", "Prof2!C21:C26"]
    // console.log("[DEBUG] rangesToRead:", rangesToRead);
    // var chosenStudentListValues = ssSource.getRangeList(rangesToRead);
  
    // var valuesToInsert = [];
    // console.assert(professorSheetNames * chosenStudentsTablesA1Notations.length == chosenStudentListValues.length);
    // var i = 0;
    // for(let professorName of professorSheetName) {
    //   for(let level in chosenStudentsTablesA1Notations) {
    //     let rowToInsert = chosenStudentListValues[i].map(studentName => [studentName, professorName, level + 1]);
    //     valuesToInsert.push(rowToInsert);
    //     i += 1;
    //   }
    // }
  
    var valuesToInsert = [];
    for(let professorSheet of professorSheets) {
      professorSheetName = professorSheet.getName();
      if (isSheetNonData(professorSheetName)) {
        Logger.log("[INFO] Found non-data sheet: '%s'. Skipping..", professorSheetName);
        continue
      }
      Logger.log("[INFO] Reading 'Chosen Students Table' from professorSheet: '%s'", professorSheetName);
  
      for(let level of Object.keys(chosenStudentsTablesA1Notations)) {
        let chosenStudentsTablesA1Notation = chosenStudentsTablesA1Notations[level];
        let rangeValues = professorSheet.getRange(chosenStudentsTablesA1Notation).getValues();
        let rowToInsert = rangeValues.filter(studentName => studentName[0] != "")
                                     .map(studentName => [_extractStudentName(studentName), professorSheetName, parseInt(level)]);
        console.log("[DEBUG] rowToInsert", rowToInsert);
        valuesToInsert.push(...rowToInsert);
      }
    }
    console.log("[DEBUG] valuesToInsert:", valuesToInsert);
  
    var ssTarget = SpreadsheetApp.openByUrl(spreadsheetDBUrl);
  
    var supervisionsSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.SUPERVISIONS);
  
    supervisionsSheet.getRange(2, 1, supervisionsSheet.getMaxRows(), supervisionsSheet.getMaxColumns()).clearContent();
    if (valuesToInsert.length > 0) {
      supervisionsSheet.getRange(2, 1, valuesToInsert.length, valuesToInsert[0].length).setValues(valuesToInsert);
    } else {
      Logger.log("[INFO] All professors' 'Chosen Student List Table' is empty. No values were inserted.")
    }
  }
  
  function _extractStudentName(studentNameWithNRP) {
    if (studentNameWithNRP != "-") {
      return studentNameWithNRP[0].split(" - ", 2)[1]
    }
    return studentNameWithNRP
  }
  
  function getSupervisionsOnLevel(supervisionSheet, level, groupByStudent=true) {
    let tableObject = getTableFromSheet(supervisionSheet, CONST.SHEET_NAMES.SUPERVISIONS);
    let data = tableObject["data"];
    let col2Idx = tableObject["col2Idx"];
  
    let studentIdx = col2Idx["Student"];
    let professorIdx = col2Idx["Professor"];
    let levelIdx = col2Idx["Level"];
  
    let supervisions = data.filter(row => row[levelIdx] == level);
    let groupColumnIdx = groupByStudent ? studentIdx : professorIdx;
    let supervisionsGrouped =  groupByColumn(supervisions, groupColumnIdx, removeColumn=true);
    return supervisionsGrouped;  // e.g. {"<student_name_1>": [[<professor_name_1>, <level>], [<professor_name_2>, <level>]], ...}
  }
  
  function getMaxStudentsBySupervisionLevel(supervisionLevelsSheet) {
    let tableObject = getTableFromSheet(supervisionLevelsSheet, CONST.SHEET_NAMES.SUPERVISION_LEVELS);
    let data = tableObject["data"];
    let col2Idx = tableObject["col2Idx"];
  
    let levelIdx = col2Idx["Level"];
    let maxStudentsIdx = col2Idx["Max Students"];
  
    let result = {};
    data.forEach(row => result[row[levelIdx]] = parseInt(row[maxStudentsIdx]));
    return result;  // e.g. {"1": 6, "2": 4, "3": 2}
  }
  
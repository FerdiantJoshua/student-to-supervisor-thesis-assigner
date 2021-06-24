function populateFormDropdowns(formTargetUrl, spreadsheetDBUrl, supervisionLevel, nameDropdownId, topicDropdownIds, professorDropdownId){
  let form = FormApp.openByUrl(formTargetUrl);
  let ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);

  let supervisionSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.SUPERVISIONS);
  let supervisionsGroupedByStudent = getSupervisionsOnLevel(supervisionSheet, supervisionLevel, groupByStudent=true);
  let studentsToOmit = Object.keys(supervisionsGroupedByStudent);

  let supervisionLevelsSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.SUPERVISION_LEVELS);
  let maxStudents = getMaxStudentsBySupervisionLevel(supervisionLevelsSheet)[supervisionLevel];
  let supervisionsGroupedByProfessor = getSupervisionsOnLevel(supervisionSheet, supervisionLevel, groupByStudent=false);
  let professorsToOmit = Object.keys(supervisionsGroupedByProfessor)
                               .filter(professorName => supervisionsGroupedByProfessor[professorName].length >= maxStudents);

  var studentSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.STUDENTS);
  _updateDropdownName(form, nameDropdownId, studentSheet, studentsToOmit);

  var topicDropdownIds = topicDropdownIds.split(",");
  var topicSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.TOPICS);
  _updateDropdownTopics(form, topicDropdownIds, topicSheet);

  var professorSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.PROFESSORS);
  _updateDropdownProfessor(form, professorDropdownId, professorSheet, professorsToOmit);
}

function _updateDropdownName(form, nameDropdownId, studentSheet, namesToOmit) {
  try {
    var nameDropdown = form.getItemById(nameDropdownId).asListItem();
  } catch {
    console.warn("[WARNING] nameDropdown with ID: %s is not found. Skipping update..", nameDropdownId)
    return
  }

  let tableObject = getTableFromSheet(studentSheet, verboseTableName=CONST.SHEET_NAMES.STUDENTS);
  let {data, col2Idx} = tableObject;
  let {"Nama": nameIdx, "NRP": nrpIdx} = col2Idx;

  let namesToOmitSet = new Set(namesToOmit);

  // Take all students which are not in "namesToOmit"
  var formItemsStudent = [];
  for(let i = 0; i < data.length; i++) {
    let row = data[i];
    if (row[nameIdx] != "" && !namesToOmitSet.delete(row[nameIdx])) {
      formItemsStudent.push(`${row[nrpIdx]} - ${row[nameIdx]}`);
    }
  }

  if (formItemsStudent.length == 0) formItemsStudent = ["All students have already got their supervisors"];
  nameDropdown.setChoiceValues(formItemsStudent);
}

function _updateDropdownTopics(form, topicDropdownIds, topicSheet) {
  let tableObject = getTableFromSheet(topicSheet, verboseTableName=CONST.SHEET_NAMES.TOPICS);
  let {data, col2Idx} = tableObject;

  for(let topicDropdownId of topicDropdownIds) {
    try {
      var topicDropdown = form.getItemById(topicDropdownId).asListItem();
    } catch {
      console.warn("[WARNING] topicDropdown with ID: %s is not found. Skipping update..", topicDropdownId)
      return
    }

    let nameIdx = col2Idx["Name"];
    let formItemsTopic = [];
    for(let row of data)
      if(row[nameIdx] != "")
        formItemsTopic.push(row[nameIdx]);
    
    topicDropdown.setChoiceValues(formItemsTopic)
  }
}

function _updateDropdownProfessor(form, professorDropdownId, professorSheet, namesToOmit) {
  try {
    var professorDropdown = form.getItemById(professorDropdownId).asListItem()
  } catch {
    console.warn("[WARNING] professorDropdown with ID: %s is not found. Skipping update..", professorDropdownId)
    return
  }

  let tableObject = getTableFromSheet(professorSheet, verboseTableName=CONST.SHEET_NAMES.PROFESSORS);
  let {data, col2Idx} = tableObject;
  let {"Nama": nameIdx, "Kelompok Keilmuan": topicIdx} = col2Idx;

  let namesToOmitSet = new Set(namesToOmit);

  // Take all professors which are not in "namesToOmit"
  var formItemsProfessor = [];
  for(let i = 0; i < data.length; i++) {
    let row = data[i];
    if (row[nameIdx] != "" && !namesToOmitSet.delete(row[nameIdx])) {
      formItemsProfessor.push(`${row[nameIdx]} - ${row[topicIdx]}`);
    }
  }
  
  if (formItemsProfessor.length == 0) formItemsProfessor = ["All professors have already reached maximum number of students"];
  professorDropdown.setChoiceValues(formItemsProfessor);
}

function generateProfessorSheets(spreadsheetDBUrl, spreadsheetAssignmentUrl) {
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);

  // Read professor table from sheet "Professors"
  let professorSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.PROFESSORS);
  let tableObject = getTableFromSheet(professorSheet, verboseTableName=CONST.SHEET_NAMES.PROFESSORS);
  let {data, col2Idx} = tableObject;
  let {"Nama": nameIdx, "Kelompok Keilmuan": topicIdx} = col2Idx;

  // Fetch "Supervision Level" table from sheet
  let supervisionLevelsSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.SUPERVISION_LEVELS);
  let maxStudentsBySupervisionLevel = getMaxStudentsBySupervisionLevel(supervisionLevelsSheet);
  // console.log("[DEBUG] supervisionsGroupedByProfessor:", supervisionsGroupedByProfessor);
  // console.log("[DEBUG] maxStudentsBySupervisionLevel:", maxStudentsBySupervisionLevel);

  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  var templateSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);

  // Fetch "Chosen Students" table props info from "Template" sheet
  var props_ChosenStudentsTables = getPropsOfChosenStudentsTables(templateSheet);
  console.log("[DEBUG] props_ChosenStudentsTables:", props_ChosenStudentsTables);
  
  // TERMINATE IF the spreadsheets has been PREPARED
  if (ssTarget.getNumSheets() > 2) {
    throw new AlreadyPreparedSpreadsheetException(
      `Spreadsheet is already prepared! If you WANT TO RESET the spreadsheet, run operation "Delete All Professor Sheets" first.`,
      debug={"spreadsheetAssignmentUrl": spreadsheetAssignmentUrl}
    );
  }

  // TERMINATE IF "Chosen Students" tables sizes are not equal to maxStudents
  for(let level of Object.keys(props_ChosenStudentsTables)) {
    let tableSize = props_ChosenStudentsTables[level].size;
    let maxStudents = maxStudentsBySupervisionLevel[level];
    if (tableSize != maxStudents)
      throw new ChosenStudentsTableSizeMismatch(
        `'Chosen Students Table' on supervision level ${level} size mismatch the maxStudents (${tableSize} vs ${maxStudents})`
      )
  }

  for(let row of data) {
    let professorName = row[nameIdx];
    if(professorName != ""){
      Logger.log("[INFO] Generating sheet for: '%s'", professorName);
      let topic = row[topicIdx];
      let newProfessorSheet = templateSheet.copyTo(ssTarget).setName(professorName);
      newProfessorSheet.getRange("D1:D2").setValues([[professorName], [topic]]);
    }
  }
}

function updateAssignmentSheetsProtection(spreadsheetDBUrl, spreadsheetAssignmentUrl, supervisionLevel) {
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);

  // Read professor table from sheet "Professors"
  let professorSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.PROFESSORS);
  let tableObject = getTableFromSheet(professorSheet, verboseTableName=CONST.SHEET_NAMES.PROFESSORS);
  let {data, col2Idx} = tableObject;
  let {"Nama": nameIdx, "Email": emailIdx} = col2Idx;

  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  // Fetch "Chosen Students" table props info from "Template" sheet
  var templateSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.TEMPLATE);
  var props_ChosenStudentsTables = getPropsOfChosenStudentsTables(templateSheet);
  console.log("[DEBUG] props_ChosenStudentsTables:", props_ChosenStudentsTables);

  // TERMINATE IF supervisionLevel IS NOT A MEMBER OF props_ChosenStudentsTables
  if (props_ChosenStudentsTables[supervisionLevel] == null) {
    throw new ChosenStudentsTableDoesNotExist(
      `'Chosen Students Table' on supervision level ${supervisionLevel} does not not exist!`,
      debug={"props_ChosenStudentsTables": props_ChosenStudentsTables}
    )
  }

  let me = Session.getEffectiveUser();
  // NOTE: this line must be executed before "_updateSheetProtection" to prevent existing editors to have access to the protected sheet
  _setUserAsOnlyEditorOfSpreadsheet(me, ssTarget);

  var emailsToAddAsEditor = []
  for(let row of data) {
    let professorName = row[nameIdx];
    if(professorName != "") {
      let professorSheet = ssTarget.getSheetByName(professorName);
      if (professorSheet == null) {
        console.warn(`[WARN]: Professor sheet with name "${professorName} is not found! Skipping protection update.."`)
        continue;
      }
      let email = row[emailIdx];
      emailsToAddAsEditor.push(email);
      
      _updateSheetProtection(professorSheet, me, email, props_ChosenStudentsTables, supervisionLevel);
    }
  }

  // manage non-professor sheets and spreadsheet protection
  let formResponsesSheet = ssTarget.getSheetByName("Form Responses");
  _setUserAsOnlyEditorOfProtection(me, formResponsesSheet.protect().setDescription('Admin Only'));
  _setUserAsOnlyEditorOfProtection(me, templateSheet.protect().setDescription('Admin Only'));
  ssTarget.addEditors(emailsToAddAsEditor); // required to let the delegated professors to edit
}

function _updateSheetProtection(professorSheet, currentUser, professorEmail, props_ChosenStudentsTables, supervisionLevel) {
  let professorName = professorSheet.getName();
  let previousSheetProtection = professorSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let previousProtections = professorSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  if (previousSheetProtection.length == 0 || previousProtections.length == 0) {
    Logger.info(`[INFO] Initializing new protections for "${professorName}" to "${professorEmail}"`);
    if (previousSheetProtection.length != 0) {
      previousSheetProtection[0].remove();
    } else if (previousProtections.length != 0) {
      console.warn(`[WARNING] Sheet "${professorName}" has already had ${previousProtections.length} protection ranges!`)
    }
    let sheetProtection = professorSheet.protect().setDescription("Admin Only").addEditor(currentUser);

    var delegatedRanges = {"ranges": [], "enabled": []};
    for(let level of Object.keys(props_ChosenStudentsTables)) {
      let props_ChosenStudentsTable = props_ChosenStudentsTables[level];
      let rangeToDelegate = professorSheet.getRange(
        props_ChosenStudentsTable.firstRow, CONST.TEMPLATE_FIRST_DATA_COLUMN,
        props_ChosenStudentsTable.size, 1
      );
      let enabled = level == supervisionLevel;
      delegatedRanges.ranges.push(rangeToDelegate);
      delegatedRanges.enabled.push(enabled);
    }

    sheetProtection.setUnprotectedRanges(delegatedRanges.ranges);
    for(let i = 0; i < delegatedRanges.ranges.length; i++) {
      let delegatedRange = delegatedRanges.ranges[i];
      let enabled = delegatedRanges.enabled[i];
      let protection = delegatedRange.protect().addEditor(currentUser);
      if (enabled) {
        protection.setDescription(`${professorEmail} Only`).addEditor(professorEmail);
      } else {
        protection.setDescription(`(Disabled) ${professorEmail} Only`);
      }
    }

  } else {
    Logger.info(`[INFO] Updating existing protections for "${professorName}" to "${professorEmail}"`);
    let currentChosenStudentTable = props_ChosenStudentsTables[supervisionLevel];
    for(let protection of previousProtections) {
      let protectionFirstCell = protection.getRange().getCell(1, 1);
      if (protectionFirstCell.getColumn() == CONST.TEMPLATE_FIRST_DATA_COLUMN
          && currentChosenStudentTable.firstRow == protectionFirstCell.getRow()) {
        protection.setDescription(`${professorEmail} Only`).addEditor(professorEmail)
      } else {
        protection.setDescription(`(Disabled) ${professorEmail} Only`).removeEditor(professorEmail)
      }
    }
  }
}

function _setUserAsOnlyEditorOfSpreadsheet(user, spreadsheet) {
  spreadsheet.addEditor(user);
  for(let user of spreadsheet.getEditors()) {
    spreadsheet.removeEditor(user);
  }
}

function _setUserAsOnlyEditorOfProtection(user, protection) {
  protection.addEditor(user);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function changeAllEditorsToViewers(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  var me = Session.getEffectiveUser();
  ssTarget.addEditor(me);
  for(let user of ssTarget.getEditors()) {
    ssTarget.removeEditor(user);
    ssTarget.addViewer(user);
  }
}

function changeAllViewersToEditors(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  for(let user of ssTarget.getViewers()) {
    ssTarget.addEditor(user);
  }
}

function removeAllEditorsAndProtections(spreadsheetAssignmentUrl) {
  let ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  ssTarget.getEditors()
          .forEach(user => ssTarget.removeEditor(user));

  let sheets = ssTarget.getSheets();
  for(let sheet of sheets) {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].remove();
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
         .forEach(range => range.remove());
  }
}

function deleteAllProfessorSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  var professorSheets = ssTarget.getSheets();
  for(var professorSheet of professorSheets) {
    if (isSheetNonData(professorSheet.getName())) {
      Logger.log("[INFO] Accessing non-data sheet: '%s'. Skipping..", professorSheet.getName())
      continue
    }
    Logger.log("[INFO] Deleting sheet: '%s'", professorSheet.getName())

    ssTarget.deleteSheet(professorSheet);
  }
  let me = Session.getEffectiveUser();
  _setUserAsOnlyEditorOfSpreadsheet(me, ssTarget);
}

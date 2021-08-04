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
  let {"NRP": nrpIdx, "Nama": nameIdx, "Is Skipped from Registration Form": isSkippedIdx} = col2Idx;

  let namesToOmitSet = new Set(namesToOmit);

  // Take all students which are not in "namesToOmit"
  var formItemsStudent = [];
  for(let i = 0; i < data.length; i++) {
    let row = data[i];
    if (row[nameIdx] != "" && !row[isSkippedIdx] && !namesToOmitSet.delete(row[nameIdx])) {
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
  let {"Nama": nameIdx, "Kelompok Keilmuan": topicIdx, "Is Skipped from Registration Form": isSkippedIdx} = col2Idx;

  let namesToOmitSet = new Set(namesToOmit);

  // Take all professors which are not in "namesToOmit"
  var formItemsProfessor = [];
  for(let i = 0; i < data.length; i++) {
    let row = data[i];
    if (row[nameIdx] != "" && !row[isSkippedIdx] && !namesToOmitSet.delete(row[nameIdx])) {
      formItemsProfessor.push(`${row[nameIdx]} - ${row[topicIdx]}`);
    }
  }
  
  if (formItemsProfessor.length == 0) formItemsProfessor = ["All professors have already reached maximum number of students"];
  professorDropdown.setChoiceValues(formItemsProfessor);
}

function setFormOpenCloseDatetime(formTargetUrl, datetimeOpen, datetimeClose) {
  let today = new Date();

  if (datetimeOpen) {
    let openDate = new Date(datetimeOpen);
    if (today < openDate) {
      let openTrigger = ScriptApp.newTrigger("_openFormTrigger")
                                .timeBased()
                                .at(openDate)
                                .create();
      setupTriggerArguments(openTrigger, [formTargetUrl], false);
    } else {
      _openForm(formTargetUrl);
    }
  }

  if (datetimeClose) {
    let closeDate = new Date(datetimeClose);
    if (today < closeDate) {
      let closeTrigger = ScriptApp.newTrigger("_closeFormTrigger")
                                  .timeBased()
                                  .at(closeDate)
                                  .create();
      setupTriggerArguments(closeTrigger, [formTargetUrl, datetimeClose], false);
    } else {
      _closeForm(formTargetUrl, today.toString().substring(4,21));
    }
  }
}

function _openFormTrigger(event) {
  var functionArgs = handleTriggered(event.triggerUid);
  Logger.log("Function arguments: %s", functionArgs);

  let [formTargetUrl] = functionArgs;
  _openForm(formTargetUrl);
}

function _openForm(formTargetUrl) {
  FormApp.openByUrl(formTargetUrl).setAcceptingResponses(true);
}

function _closeFormTrigger(event) {
  var functionArgs = handleTriggered(event.triggerUid);
  Logger.log("Function arguments: %s", functionArgs);

  let [formTargetUrl, datetimeClose] = functionArgs;
  _closeForm(formTargetUrl, datetimeClose);
}

function _closeForm(formTargetUrl, datetimeClose) {
  FormApp.openByUrl(formTargetUrl)
         .setCustomClosedFormMessage(`Formulir telah ditutup pada "${datetimeClose}".`)
         .setAcceptingResponses(false);
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
  if (ssTarget.getSheets().filter(sheet => isSheetNonData(sheet.getSheetName())).length == 0) {
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

function updateAssignmentSheetsProtection(spreadsheetDBUrl, spreadsheetAssignmentUrl, supervisionLevel, nDataSkipped=0) {
  var startTime = new Date().getTime();
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

  // Protect non-professor sheets protection
  let manualSheet = ssTarget.getSheetByName(CONST.SHEET_NAMES.MANUAL_BOOK);
  _setUserAsOnlyEditorOfProtection(me, manualSheet.protect().setDescription('Admin Only'));
  _setUserAsOnlyEditorOfProtection(me, templateSheet.protect().setDescription('Admin Only'));

  var emailsToAddAsEditor = []
  Logger.log(`[INFO] Skipping ${nDataSkipped} data as requested`);
  for(var i = nDataSkipped; i < data.length; i++) {
    let row = data[i];
    let professorName = row[nameIdx];
    if(professorName != "") {
      // Self-terminate on internal timeout with email notification to let user continue the operation
      if ((new Date().getTime() - startTime) / 1000 >= CONST.APP_TIMEOUT_SECONDS){
        emailReportToSelf(
          subject="Partially Complete Operation Notification",
          message=`Operation "updateAssignmentSheetsProtection" was stopped due to timeout.\n\n`
          + `Please rerun the operation with this parameter "Amount of Data to be Skipped = ${i}"`
        );
        throw new PartiallyExecutedOperationException();
      }

      let professorSheet = ssTarget.getSheetByName(professorName);
      if (professorSheet == null) {
        console.warn(`[WARNING]: Professor sheet with name "${professorName} is not found! Skipping protection update.."`)
        continue;
      }
      let email = row[emailIdx];
      emailsToAddAsEditor.push(email);
      
      // INITIALIZE or UPDATE protections
      let prevSheetProt = professorSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      let prevProts = professorSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      if (prevSheetProt.length == 0 || prevProts.length == 0) {
        Logger.info(`[INFO] Initializing new protections for "${professorName}" to "${email}"`);
        _initializeSheetProtection(
          professorSheet, me, email, props_ChosenStudentsTables, supervisionLevel, prevSheetProt, prevProts
        )
      } else {
        // Fetch "Supervisions" table from sheet
        let supervisionSheet = ssSource.getSheetByName(CONST.SHEET_NAMES.SUPERVISIONS);
        let supervisionsGroupedByProfessor = getSupervisionsOnLevel(supervisionSheet, supervisionLevel, groupByStudent=false);
        let prevStudents = supervisionsGroupedByProfessor[professorName];
        let nPrevStudents = prevStudents == null ? 0 : prevStudents.length;
        Logger.info(`[INFO] Updating existing protections for "${professorName}" to "${email}"`);
        _updateSheetProtection(
          professorSheet, email, props_ChosenStudentsTables, supervisionLevel, nPrevStudents, prevSheetProt, prevProts
        );
      }
    }
  }
  
  // Invite professors to view spreadsheet by their email
  ssTarget.addViewers(emailsToAddAsEditor);
}

function _initializeSheetProtection(
  professorSheet, currentUser, professorEmail, props_ChosenStudentsTables, supervisionLevel, prevSheetProt, prevProts
) {
  if (prevSheetProt.length != 0) {
    prevSheetProt[0].remove();
  } else if (prevProts.length != 0) {
    prevProts.forEach(previousProtection => previousProtection.remove());
  }
  let sheetProtection = professorSheet.protect().setDescription("Admin Only").addEditor(currentUser);

  // Create ranges to be unprotected (delegated to professor)
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

  // Unprotect and delegate the range
  sheetProtection.setUnprotectedRanges(delegatedRanges.ranges);
  for(let i = 0; i < delegatedRanges.ranges.length; i++) {
    let delegatedRange = delegatedRanges.ranges[i];
    let enabled = delegatedRanges.enabled[i];
    let protection = delegatedRange.protect().addEditor(currentUser);
    if (enabled) {
      protection.setDescription(`${professorEmail} Only`).addEditor(professorEmail);
      delegatedRange.setBackground("yellow");
    } else {
      protection.setDescription(`(Disabled) ${professorEmail} Only`);
      protection.getRange().setBackground("#B7B7B7"); // dark gray 1
    }
  }
}

function _updateSheetProtection(
  professorSheet, professorEmail, props_ChosenStudentsTables, supervisionLevel, nRowsFilled=0, prevSheetProt, prevProts
) {
  let currentChosenStudentTable = props_ChosenStudentsTables[supervisionLevel];

  // Enable or Disable professor's protection ranges
  for(let protection of prevProts) {
    let protectionFirstCell = protection.getRange().getCell(1, 1);
    if (protectionFirstCell.getColumn() == CONST.TEMPLATE_FIRST_DATA_COLUMN
        && currentChosenStudentTable.firstRow == protectionFirstCell.getRow()) {
      protection.setDescription(`${professorEmail} Only`).addEditor(professorEmail);
      protection.getRange().setBackground("#00FF00"); // green; empty or unsaved ranges will be "yellowed" below
    } else {
      protection.setDescription(`(Disabled) ${professorEmail} Only`).removeEditor(professorEmail);
      protection.getRange().setBackground("#B7B7B7"); // dark gray 1
    }
  }

  // Adjust unprotected ranges to remaining slots (based on nRowsFilled)
  let newUnprotectedRanges = [];
  for(let unprotectedRange of prevSheetProt[0].getUnprotectedRanges()) {
    if (_isRangeAChosenStudentTable(unprotectedRange, currentChosenStudentTable)) {
      let newUnprotectedRange = nRowsFilled > 0 && currentChosenStudentTable.size <= nRowsFilled ? null : professorSheet.getRange(
        currentChosenStudentTable.firstRow + nRowsFilled, unprotectedRange.getCell(1, 1).getColumn(),
        currentChosenStudentTable.size - nRowsFilled, unprotectedRange.getNumColumns()
      );
      if (newUnprotectedRange != null) {
        newUnprotectedRanges.push(newUnprotectedRange);
        newUnprotectedRange.setBackground("yellow");
      }
    } else {
      newUnprotectedRanges.push(unprotectedRange);
    }
  }
  prevSheetProt[0].setUnprotectedRanges(newUnprotectedRanges);
}

function _isRangeAChosenStudentTable(protectionRange, props_chosenStudentTable) {
  let firstCell = protectionRange.getCell(1, 1);
  let firstRow = firstCell.getRow();
  // console.log(`[DEBUG] firstrow = ${firstRow}, props_chosenStudentTable =`, props_chosenStudentTable);

  return firstCell.getColumn() == CONST.TEMPLATE_FIRST_DATA_COLUMN
          && firstRow >= props_chosenStudentTable.firstRow
          && firstRow < props_chosenStudentTable.firstRow + props_chosenStudentTable.size;
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

function setAssignmentSpreadsheetGrantRevokeDatetime(spreadsheetAssignmentUrl, datetimeGrant, datetimeRevoke) {
  let today = new Date();

  if (datetimeGrant) {
    let grantDate = new Date(datetimeGrant);
    if (today < grantDate) {
      let grantTrigger = ScriptApp.newTrigger("_grantAccessTrigger")
                                .timeBased()
                                .at(grantDate)
                                .create();
      setupTriggerArguments(grantTrigger, [spreadsheetAssignmentUrl], false);
    } else {
      _changeAllViewersToEditors(spreadsheetAssignmentUrl);
    }
  }

  if (datetimeRevoke) {
    let revokeDate = new Date(datetimeRevoke);
    if (today < revokeDate) {
      let revokeTrigger = ScriptApp.newTrigger("_revokeAccessTrigger")
                                  .timeBased()
                                  .at(revokeDate)
                                  .create();
      setupTriggerArguments(revokeTrigger, [spreadsheetAssignmentUrl], false);
    } else {
      _changeAllEditorsToViewers(spreadsheetAssignmentUrl);
    }
  }
}

function _grantAccessTrigger(event) {
  var functionArgs = handleTriggered(event.triggerUid);
  let [spreadsheetAssignmentUrl] = functionArgs;
  _changeAllViewersToEditors(spreadsheetAssignmentUrl);
}

function _revokeAccessTrigger(event) {
  var functionArgs = handleTriggered(event.triggerUid);
  let [spreadsheetAssignmentUrl] = functionArgs;
  _changeAllEditorsToViewers(spreadsheetAssignmentUrl);
}

function _changeAllViewersToEditors(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  for(let user of ssTarget.getViewers()) {
    ssTarget.addEditor(user);
  }
}

function _changeAllEditorsToViewers(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  var me = Session.getEffectiveUser();
  ssTarget.addEditor(me);
  for(let user of ssTarget.getEditors()) {
    ssTarget.removeEditor(user);
    ssTarget.addViewer(user);
  }
}

function removeAllEditorsAndProtections(spreadsheetAssignmentUrl) {
  let ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  ssTarget.getEditors()
          .forEach(user => ssTarget.removeEditor(user));

  let sheets = ssTarget.getSheets();
  for(let sheet of sheets) {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
         .forEach(protection => protection.remove());
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
         .forEach(protection => {protection.getRange().setBackground("yellow"); protection.remove();});
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

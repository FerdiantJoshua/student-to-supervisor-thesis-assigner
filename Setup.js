// var formTarget = "1FBWSBfQhAif5CE105G13HPGwekNckZjzMh9eleLa28M";
// var spreadsheetSourceId = "1bhuvOg6tJOLQyvvyeXDEAUaGwG_dfgavuY6j0fWRjeg";
// var spreadsheetTargetId = "1TIiC1p0vMptRFfcFISsTG6sJx-W_1Fo8cPD2FJTGv5s";

function populateFormDropdowns(formTargetUrl, spreadsheetDBUrl, nameDropdownId, topicDropdownIds, professorDropdownId){
  var form = FormApp.openByUrl(formTargetUrl);
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);

  var studentSheet = ssSource.getSheetByName("Students");
  _updateDropdownName(form, nameDropdownId, studentSheet);

  // var topicDropdownIds = ["2083436530", "25904452", "313624459"];
  var topicDropdownIds = topicDropdownIds.split(",");
  var topicSheet = ssSource.getSheetByName("Topics");
  _updateDropdownTopics(form, topicDropdownIds, topicSheet);

  var professorSheet = ssSource.getSheetByName("Professors");
  _updateDropdownProfessor(form, professorDropdownId, professorSheet);
}

function _updateDropdownName(form, nameDropdownId, studentSheet) {
  try {
    var nameDropdown = form.getItemById(nameDropdownId).asListItem();
  } catch {
    console.warn("[WARNING] nameDropdown with ID: %s is not found. Skipping update..")
    return
  }

  // grab the values in the first column of the sheet - use 2 to skip header row
  var nrpValues = studentSheet.getRange(2, 2, studentSheet.getMaxRows()).getValues();
  var nameValues = studentSheet.getRange(2, 3, studentSheet.getMaxRows()).getValues();

  var formItemStudent = [];

  for(var i = 0; i < nrpValues.length; i++)   
    if(nrpValues[i][0] != "")
      formItemStudent[i] = nrpValues[i][0] + " - " + nameValues[i][0];

  nameDropdown.setChoiceValues(formItemStudent);

}

function _updateDropdownTopics(form, topicDropdownIds, topicSheet) {
  var topicValues = topicSheet.getRange(2, 2, topicSheet.getMaxRows()).getValues();
  for(var topicDropdownId of topicDropdownIds) {
    try {
      var topicDropdown = form.getItemById(topicDropdownId).asListItem();
    } catch {
      console.warn("[WARNING] topicDropdown with ID: %s is not found. Skipping update..")
      return
    }

    var formItemTopic = [];

    for(var i = 0; i < topicValues.length; i ++)
      if(topicValues[i][0] != "")
        formItemTopic[i] = topicValues[i][0]
    
    topicDropdown.setChoiceValues(formItemTopic)
  }
}

function _updateDropdownProfessor(form, professorDropdownId, professorSheet) {
  try {
    var professorDropdown = form.getItemById(professorDropdownId).asListItem()
  } catch {
    console.warn("[WARNING] professorDropdown with ID: %s is not found. Skipping update..")
    return
  }
  
  var nameValues = professorSheet.getRange(2, 3, professorSheet.getMaxRows()).getValues();
  var topicValues = professorSheet.getRange(2, 4, professorSheet.getMaxRows()).getValues();

  var formItemProfessor = [];

  for(var i = 0; i < nameValues.length; i ++)
    if(nameValues[i][0] != "")
      formItemProfessor[i] = nameValues[i][0] + " - " + topicValues[i][0];
  
  professorDropdown.setChoiceValues(formItemProfessor);
}

function prepareResponseSheet(spreadsheetDBUrl, spreadsheetAssignmentUrl) {
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetDBUrl);

  var professorSheet = ssSource.getSheetByName("Professors");
  var nameValues = professorSheet.getRange(2, 3, professorSheet.getMaxRows()).getValues();
  var topicValues = professorSheet.getRange(2, 4, professorSheet.getMaxRows()).getValues();
  var emailValues = professorSheet.getRange(2, 5, professorSheet.getMaxRows()).getValues();


  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  if (ssTarget.getNumSheets() > 2) {
    throw new AlreadyPreparedSpreadsheetException(
      `Spreadsheet is already prepared! If you WANT TO RESET the spreadsheet, run operation "Delete All Professor Sheets" first.`,
      {"spreadsheetAssignmentUrl": spreadsheetAssignmentUrl}
    );
  }
  _setMeAsOnlyEditorOfSpreadsheet(ssTarget);

  var templateSheet = ssTarget.getSheetByName("Template");
  var emailsToAddAsEditor = []

  for(var i = 0; i < nameValues.length; i++){
    if(nameValues[i][0] != ""){
      Logger.log("[INFO]: Generating sheet for: '%s'", nameValues[i][0]);
      let professorName = nameValues[i][0];
      let topic = topicValues[i][0];
      let email = emailValues[i][0];
      emailsToAddAsEditor.push(email);
      let newProfessorSheet = templateSheet.copyTo(ssTarget).setName(professorName);
      newProfessorSheet.getRange("C1:C2").setValues([[professorName], [topic]]);
      
      _protectAndDelegateSheet(newProfessorSheet, email);
    }
  }

  // manage non-professor sheets and spreadsheets protection
  let formResponsesSheet = ssTarget.getSheetByName("Form Responses");
  _setMeAsOnlyEditorOfProtectionRange(formResponsesSheet.protect().setDescription('Admin Only'));
  _setMeAsOnlyEditorOfProtectionRange(templateSheet.protect().setDescription('Admin Only'));
  ssTarget.addEditors(emailsToAddAsEditor); // required to let the delegated professors to edit
}

function _setMeAsOnlyEditorOfSpreadsheet(spreadsheet) {
  var me = Session.getEffectiveUser();
  spreadsheet.addEditor(me);
  for(let user of spreadsheet.getEditors()) {
    spreadsheet.removeEditor(user);
  }
}

function _protectAndDelegateSheet(professorSheet, professorEmail) {
  console.log(professorEmail);
  let protection = professorSheet.protect().setDescription("Admin Only");

  // create protection range
  var delegatedRanges = [];
  for(var firstCell_ChosenStudentList of firstCell_ChosenStudentLists) {
    let topLeftCell = firstCell_ChosenStudentList.getNextNCols(-1);
    let botRightCell = topLeftCell.getNextEmptyRow(professorSheet).getNextNCols(1).getNextNRows(-2);
    console.log("topLeftCell, botRightCell: (%s, %s)", topLeftCell.getPos(), botRightCell.getPos());
    let range = professorSheet.getRange(topLeftCell.getNextNCols(1).getPos() + ":" + botRightCell.getPos());
    delegatedRanges.push(range);
  }
  protection.setUnprotectedRanges(delegatedRanges);

  _setMeAsOnlyEditorOfProtectionRange(protection);

  // add professor as the editor of delegated ranges
  for(let delegatedRange of delegatedRanges) {
    let delegatedProtection = delegatedRange.protect().setDescription(`${professorEmail} Only`);
    delegatedProtection.addEditor(professorEmail);
    // delegatedProtection.addViewer(professorEmail);
  }
}

function _setMeAsOnlyEditorOfProtectionRange(protection) {
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function escalateAllUsersAccess(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  for(let user of ssTarget.getViewers()) {
    ssTarget.addEditor(user);
  }
}

function deescalateAllUsersAccess(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  var me = Session.getEffectiveUser();
  ssTarget.addEditor(me);
  for(let user of ssTarget.getEditors()) {
    ssTarget.removeEditor(user);
    ssTarget.addViewer(user);
  }
}

function deleteAllProfessorSheets(spreadsheetAssignmentUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);

  var professorSheets = ssTarget.getSheets().slice(2);
  for(var professorSheet of professorSheets) {
    Logger.log("[INFO]: Deleting sheet: '%s'", professorSheet.getName())
    
    let loweredSheetName = professorSheet.getName().toLowerCase();
    if (loweredSheetName == "template" || loweredSheetName.includes("form responses")) {
      Logger.log("[INFO]: Accessing non-data sheet: '%s'. Skipping..", professorSheet.getName())
      continue
    }

    ssTarget.deleteSheet(professorSheet);
  }
  _setMeAsOnlyEditorOfSpreadsheet(ssTarget);
}

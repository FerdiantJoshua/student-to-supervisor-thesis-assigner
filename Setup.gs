// var formTarget = "1FBWSBfQhAif5CE105G13HPGwekNckZjzMh9eleLa28M";
// var spreadsheetSourceId = "1bhuvOg6tJOLQyvvyeXDEAUaGwG_dfgavuY6j0fWRjeg";
// var spreadsheetTargetId = "1TIiC1p0vMptRFfcFISsTG6sJx-W_1Fo8cPD2FJTGv5s";

function updateForm(formTargetUrl, spreadsheetSourceUrl, nameDropdownId, topicDropdownIds, professorDropdownIds){
  var form = FormApp.openByUrl(formTargetUrl);
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetSourceUrl);

  var nameDropdown = form.getItemById(nameDropdownId).asListItem();
  var studentData = ssSource.getSheetByName("Students");
  updateNameDropdown(nameDropdown, studentData);

  // var topicDropdownIds = ["2083436530", "25904452", "313624459"];
  var topicDropdownIds = topicDropdownIds.split(",");
  var topicData = ssSource.getSheetByName("Topics");
  for(var topicDropdownId of topicDropdownIds)
    updateDropdownTopic(form.getItemById(topicDropdownId).asListItem(), topicData);

  // var professorDropdownIds = ["1187134166", "103959848", "1540995051"];
  var professorDropdownIds = professorDropdownIds.split(",");
  var professorData = ssSource.getSheetByName("Professors");
  for(var professorDropdownId of professorDropdownIds)
    updateDropdownProfessor(form.getItemById(professorDropdownId).asListItem(), professorData);
}

function updateNameDropdown(nameDropdown, studentData) {
  // grab the values in the first column of the sheet - use 2 to skip header row
  var nrpValues = studentData.getRange(2, 2, studentData.getMaxRows()).getValues();
  var nameValues = studentData.getRange(2, 3, studentData.getMaxRows()).getValues();

  var formItemStudent = [];

  for(var i = 0; i < nrpValues.length; i++)   
    if(nrpValues[i][0] != "")
      formItemStudent[i] = nrpValues[i][0] + " - " + nameValues[i][0];

  nameDropdown.setChoiceValues(formItemStudent);

}

function updateDropdownTopic(topicDropdown, topicData) {
  // grab the values in the first column of the sheet - use 2 to skip header row
  var topicValues = topicData.getRange(2, 2, topicData.getMaxRows()).getValues();

  var formItemTopic = [];

  for(var i = 0; i < topicValues.length; i ++)
    if(topicValues[i][0] != "")
      formItemTopic[i] = topicValues[i][0]
  
  topicDropdown.setChoiceValues(formItemTopic)
}

function updateDropdownProfessor(professorDropdown, professorData) {
  var nameValues = professorData.getRange(2, 3, professorData.getMaxRows()).getValues();
  var topicValues = professorData.getRange(2, 4, professorData.getMaxRows()).getValues();

  var formItemProfessor = [];

  for(var i = 0; i < nameValues.length; i ++)
    if(nameValues[i][0] != "")
      formItemProfessor[i] = nameValues[i][0] + " - " + topicValues[i][0];
  
  professorDropdown.setChoiceValues(formItemProfessor);
}

function generateProfessorSheets(spreadsheetSourceUrl, spreadsheetTargetUrl) {
  var ssSource = SpreadsheetApp.openByUrl(spreadsheetSourceUrl);

  var professorData = ssSource.getSheetByName("Professors");
  var nameValues = professorData.getRange(2, 3, professorData.getMaxRows()).getValues();
  var topicValues = professorData.getRange(2, 4, professorData.getMaxRows()).getValues();
  var emailValues = professorData.getRange(2, 5, professorData.getMaxRows()).getValues();


  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetTargetUrl);

  var templateSheet = ssTarget.getSheetByName("Template");

  for(var i = 0; i < nameValues.length; i++){
    if(nameValues[i][0] != ""){
      Logger.log("[INFO]: Generating sheet for: '%s'", nameValues[i][0]);
      let professorName = nameValues[i];
      let topic = topicValues[i];
      let email = emailValues[i];
      templateSheet.copyTo(ssTarget).setName(professorName);
      ssTarget.getSheetByName(professorName).getRange("C1:C2").setValues([[professorName], [topic]]);
    }
  }
}

function deleteAllProfessorSheets(spreadsheetTargetUrl) {
  var ssTarget = SpreadsheetApp.openByUrl(spreadsheetTargetUrl);

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
}

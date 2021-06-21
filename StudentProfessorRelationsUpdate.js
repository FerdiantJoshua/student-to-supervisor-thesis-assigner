// "ssSource" and "ssTarget" VALUES ARE REVERSED in this function!
function updateStudentProfessorRelations(spreadsheetDBUrl, spreadsheetAssignmentUrl) {
    var ssTarget = SpreadsheetApp.openByUrl(spreadsheetDBUrl);
  
    var professorTableSheet = ssTarget.getSheetByName("Professors");
    let professorIdx2Col = professorTableSheet.getRange(1, 1, 1).getValues();
    var professorCol2Idx = getCol2Idx(professorIdx2Col);
    Logger.log("[INFO]: professorCol2Idx is:\n%s", professorCol2Idx)
    var professorNames = professorTableSheet.getRange(2, professorCol2Idx["Nama"], professorTableSheet.getMaxRows()).getValues();
  
    var studentTableSheet = ssTarget.getSheetByName("Students");
    let studentIdx2Col = studentTableSheet.getRange(1, 1, 1).getValues();
    var studentCol2Idx = getCol2Idx(studentIdx2Col);
    Logger.log("[INFO]: studentCol2Idx is:\n%s", studentCol2Idx)
    var studentNames = studentTableSheet.getRange(2, studentCol2Idx["Nama"], studentTableSheet.getMaxRows()).getValues();
  
  
    var ssSource = SpreadsheetApp.openByUrl(spreadsheetAssignmentUrl);
  
    var professorSheets = ssSource.getSheets();
    if (professorSheets.length == 0)
      Logger.log("[INFO] No professor sheets to clear");
  
    for(var professorSheet of professorSheets) {
      Logger.log("[INFO] Reading chosen students list from professorSheet: '%s'", professorSheet.getName())
  
      if (_isSheetNonData(professorSheet.getName())) {
        Logger.log("[INFO] Accessing non-data sheet: '%s'. Skipping..", professorSheet.getName())
        continue
      }
    }
  
  
  }
  
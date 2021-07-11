function doGet(e) {  
  var template = HtmlService.createTemplateFromFile("index");
  template.serviceUrl = ScriptApp.getService().getUrl();
  template.batchName = CONST.HTML_DEFAULT_VALUES.BATCH_NAME;
  template.spreadsheetDatabaseUrl = CONST.HTML_DEFAULT_VALUES.SPREADSHEET_DB_URL;
  template.status = "-";
  template.statusMessage = "No message";
  var html = template.evaluate();
  return html.setTitle("Admin Page (Student-to-Supervisor Assignment)");
}

function doPost(e) {
  var errorMessage = "";
  var params = e.parameter;
  console.log("[DEBUG] params is:\n", params);

  try {
    operationParams = getOperationParametersOnBatch(params.spreadsheetDatabaseUrl, params.batchName);
    console.log("[DEBUG] operationParams is:\n", operationParams);

    let lock = LockService.getScriptLock();
    let success = lock.tryLock(CONST.LOCK_WAIT_SECONDS * 1000);
    if (!success) {
      throw new WaitLockTimeoutException(`Another operation is still running! Please wait until all operation executions complete.`)
    }

    var formResponsesSheetName = params.batchName;
    Logger.log("[INFO] Running operation: %s", params.operationType);
    switch (params.operationType) {
      // Setup
      case "populateFormDropdowns":
        populateFormDropdowns(
          operationParams["Form URL"],
          params.spreadsheetDatabaseUrl,
          operationParams["Supervision Level"],
          operationParams["Form Name Dropdown Id"],
          operationParams["Form Topic Dropdown Ids"],
          operationParams["Form Professor Dropdown Id"]
        );
        break;
      case "setFormOpenCloseDatetime":
        setFormOpenCloseDatetime(
          operationParams["Form URL"],
          params.datetimeOpen,
          params.datetimeClose,
        )
        break;
      case "generateProfessorSheets":
        generateProfessorSheets(
          params.spreadsheetDatabaseUrl,
          operationParams["Spreadsheet Assignment URL"],
        );
        break;
      case "updateAssignmentSheetsProtection":
        updateAssignmentSheetsProtection(
          params.spreadsheetDatabaseUrl,
          operationParams["Spreadsheet Assignment URL"],
          operationParams["Supervision Level"],
          parseInt(params.nDataSkipped),
        );
        break;
      case "setAssignmentSpreadsheetGrantRevokeDatetime":
        setAssignmentSpreadsheetGrantRevokeDatetime(
          operationParams["Spreadsheet Assignment URL"],
          params.datetimeOpen,
          params.datetimeClose,
          );
        break;
      case "removeAllEditorsAndProtections":
        removeAllEditorsAndProtections(operationParams["Spreadsheet Assignment URL"]);
        break;
      case "deleteAllProfessorSheets":
        deleteAllProfessorSheets(operationParams["Spreadsheet Assignment URL"]);
        break;
      // ResponseManagement
      case "assignStudentsToSheets":
        assignStudentsToSheets(
          params.spreadsheetDatabaseUrl,
          operationParams["Spreadsheet Assignment URL"],
          operationParams["Spreadsheet Form Responses URL"],
          formResponsesSheetName,
        );
        break;
      case "clearAllStudentQueues":
        clearAllStudentQueues(operationParams["Spreadsheet Assignment URL"]);
        break;
      // SupervisionRelationsManagement
      case "saveStudentProfessorRelations":
        saveStudentProfessorRelations(
          params.spreadsheetDatabaseUrl,
          operationParams["Spreadsheet Assignment URL"],
          params.forceSave === 'true',
        );
        break;
      default:
        throw new UnknownOperationException(`Unknown operation_type "${params.operationType}"`);
    }
  } catch(err) {
    if (isExceptionCustom(err)) {
      if (err instanceof PartiallyExecutedOperationException) {
        console.warn(`[WARNING] Partially executed operation for "${params.operationDisplayName}"!`);
        errorMessage = "WARNING";
      } else {
        console.error(err);
        errorMessage = err.message;
      }
    } else {
      throw err;
    }
  } finally {
    Logger.log("[INFO] Operation '%s' finished", params.operationType);
  }

  var template = HtmlService.createTemplateFromFile("index");
  template.serviceUrl = ScriptApp.getService().getUrl();
  template.batchName = params.batchName == CONST.BATCH_NAME_DEFAULT ? CONST.HTML_DEFAULT_VALUES.BATCH_NAME : params.batchName;
  template.spreadsheetDatabaseUrl = params.spreadsheetDatabaseUrl;
  if (errorMessage == "WARNING") {
    template.status = "WARNING";
    template.statusMessage = `Operation "${params.operationDisplayName}" partially executed! Please check your email for further action!`;
  } else if (errorMessage != "") {
    template.status = "ERROR";
    template.statusMessage = `Operation "${params.operationDisplayName}" execution failed! ${errorMessage}`;
  } else {
    template.status = "OK";
    template.statusMessage = `Operation "${params.operationDisplayName}" executed successfully!`;
  }
  var html = template.evaluate();
  
  return html.setTitle("Admin Page (Student-to-Supervisor Assignment)")
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

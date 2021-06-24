function doGet(e) {  
  var template = HtmlService.createTemplateFromFile("index");
  template.serviceUrl = ScriptApp.getService().getUrl();
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
    operationParams = getOperationParametersOnBatch(params.spreadsheet_database_url, params.batch_name);
    console.log("[DEBUG] operationParams is:\n", operationParams);

    Logger.log("[INFO] Running operation: %s", params.operation_type);
    switch (params.operation_type) {
      // Setup
      case "populateFormDropdowns":
        populateFormDropdowns(
          operationParams["Form URL"],
          params.spreadsheet_database_url,
          operationParams["Supervision Level"],
          operationParams["Form Name Dropdown Id"],
          operationParams["Form Topic Dropdown Ids"],
          operationParams["Form Professor Dropdown Id"]
        );
        break;
      case "generateProfessorSheets":
        generateProfessorSheets(
          params.spreadsheet_database_url,
          operationParams["Spreadsheet Assignment URL"],
        );
        break;
      case "updateAssignmentSheetsProtection":
        updateAssignmentSheetsProtection(
          params.spreadsheet_database_url,
          operationParams["Spreadsheet Assignment URL"],
          operationParams["Supervision Level"],
        );
        break;
      case "changeAllEditorsToViewers":
        changeAllEditorsToViewers(operationParams["Spreadsheet Assignment URL"]);
        break;
      case "changeAllViewersToEditors":
        changeAllViewersToEditors(operationParams["Spreadsheet Assignment URL"]);
        break;
      case "removeAllEditorsAndProtections":
        removeAllEditorsAndProtections(operationParams["Spreadsheet Assignment URL"]);
        break;
      case "deleteAllProfessorSheets":
        deleteAllProfessorSheets(operationParams["Spreadsheet Assignment URL"]);
        break;
      // ResponseManagement
      case "assignStudentsToSheets":
        assignStudentsToSheets(operationParams["Spreadsheet Assignment URL"]);
        break;
      case "clearAllProfessorSheets":
        clearAllProfessorSheets(operationParams["Spreadsheet Assignment URL"]);
        break;
      // SupervisionRelationsManagement
      case "saveStudentProfessorRelations":
        saveStudentProfessorRelations(
          params.spreadsheet_database_url,
          operationParams["Spreadsheet Assignment URL"]
        );
        break;
      default:
        throw new UnknownOperationException(`Unknown operation_type "${params.operation_type}"`);
    }
  } catch(err) {
    if (_isExceptionCustom(err)) {
      console.error(err);
      errorMessage = err.message;
    } else {
      throw err;
    }
  } finally {
    Logger.log("[INFO] Operation '%s' finished", params.operation_type);
  }

  var template = HtmlService.createTemplateFromFile("index");
  template.serviceUrl = ScriptApp.getService().getUrl();
  if (errorMessage == "") {
    template.status = "OK";
    template.statusMessage = "Operation executed successfully!";
  } else {
    template.status = "ERROR";
    template.statusMessage = `Operation execution failed! ${errorMessage}`;
  }
  var html = template.evaluate();
  
  return html.setTitle("Admin Page (Student-to-Supervisor Assignment)")
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

function UnknownOperationException(message) {
  this.message = message;
  this.name = 'UnknownOperationException';
}

function BatchNotFoundException(message) {
  this.message = message;
  this.name = 'BatchNotFoundException';
}

function AlreadyPreparedSpreadsheetException(message, debug=null) {
  this.message = message;
  this.name = 'AlreadyPreparedSpreadsheetException';
  this.debug = debug;
}

function UnpreparedSpreadsheetException(message, debug=null) {
  this.message = message;
  this.name = 'UnpreparedSpreadsheetException';
  this.debug = debug;
}

function StudentsAlreadyAssignedException(message, debug=null) {
  this.message = message;
  this.name = 'StudentsAlreadyAssignedException';
  this.debug = debug;
}

function ChosenStudentsTableSizeMismatch(message, debug=null) {
  this.message = message;
  this.name = 'ChosenStudentsTableSizeMismatch';
  this.debug = debug;
}

function ChosenStudentsTableDoesNotExist(message, debug=null) {
  this.message = message;
  this.name = 'ChosenStudentsTableDoesNotExist';
  this.debug = debug;
}

function _isExceptionCustom(err) {
  return err instanceof UnknownOperationException ||
    err instanceof BatchNotFoundException ||
    err instanceof AlreadyPreparedSpreadsheetException ||
    err instanceof UnpreparedSpreadsheetException ||
    err instanceof StudentsAlreadyAssignedException ||
    err instanceof ChosenStudentsTableSizeMismatch ||
    err instanceof ChosenStudentsTableDoesNotExist
}

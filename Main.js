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

    var formResponsesSheetName = params.batch_name;
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
      case "setFormOpenCloseDatetime":
        setFormOpenCloseDatetime(
          operationParams["Form URL"],
          params.datetime_open,
          params.datetime_close,
        )
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
          parseInt(params.n_data_skipped),
        );
        break;
      case "setAssignmentSpreadsheetGrantRevokeDatetime":
        setAssignmentSpreadsheetGrantRevokeDatetime(
          operationParams["Spreadsheet Assignment URL"],
          params.datetime_open,
          params.datetime_close,
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
          params.spreadsheet_database_url,
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
          params.spreadsheet_database_url,
          operationParams["Spreadsheet Assignment URL"],
          params.force_save === 'true',
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
    template.statusMessage = `Operation "${params.operation_type}" executed successfully!`;
  } else {
    template.status = "ERROR";
    template.statusMessage = `Operation "${params.operation_type}" execution failed! ${errorMessage}`;
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

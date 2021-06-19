function UnknownOperationException(message) {
  this.message = message;
  this.name = 'UnknownOperationException';
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

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

  operationParams = getOperationParametersOnBatch(params.spreadsheet_database_url, params.batch_name);
  console.log("[DEBUG] operationParams is:\n", operationParams);

  Logger.log("[INFO] Running operation: %s", params.operation_type);
  try {
    switch (params.operation_type) {
      case "updateForm":
        updateForm(
          operationParams["Form URL"],
          params.spreadsheet_database_url,
          operationParams["Form Name Dropdown Id"],
          operationParams["Form Topic Dropdown Ids"],
          operationParams["Form Professor Dropdown Ids"]
        );
        break;
      case "generateProfessorSheets":
        generateProfessorSheets(params.spreadsheet_database_url, operationParams["Spreadsheet Form Responses URL"]);
        break;
      case "deleteAllProfessorSheets":
        deleteAllProfessorSheets(operationParams["Spreadsheet Form Responses URL"]);
        break;
      case "assignStudentsToSheets":
        assignStudentsToSheets(operationParams["Spreadsheet Form Responses URL"]);
        break;
      case "clearAllProfessorSheets":
        clearAllProfessorSheets(operationParams["Spreadsheet Form Responses URL"]);
        break;
      default:
        throw new UnknownOperationException(`Unknown operation_type "${params.operation_type}"`);
    }
  } catch(err) {
    if (err instanceof UnknownOperationException) {
      console.error(err);
      errorMessage = err;
    } else {
      throw err;
    }
  }

  var template = HtmlService.createTemplateFromFile("index");
  template.serviceUrl = ScriptApp.getService().getUrl();
  if (errorMessage == "") {
    template.status = "OK";
    template.statusMessage = "Operation executed successfully!";
  } else {
    template.status = "ERROR";
    template.statusMessage = `Operation failed! Detail: ${errorMessage}`;
  }
  var html = template.evaluate();
  
  return html.setTitle("Admin Page (Student-to-Supervisor Assignment)")
}

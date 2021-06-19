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
    return html.setTitle("Pendaftaran Dosbing Data Controller");
  }
  
  function doPost(e) {
    var errorMessage = "";
    var params = e.parameter;
    console.log(params);
  
    Logger.log("[INFO] Running operation: %s", params.operation_type);
    try {
      switch (params.operation_type) {
        case "updateForm":
          updateForm(
            params.form_target_url,
            params.spreadsheet_source_url,
            params.name_dropdown_id,
            params.topic_dropdown_ids,
            params.professor_dropdown_ids
          );
          break;
        case "generateProfessorSheets":
          generateProfessorSheets(params.spreadsheet_source_url, params.spreadsheet_target_url);
          break;
        case "deleteAllProfessorSheets":
          deleteAllProfessorSheets(params.spreadsheet_target_url);
          break;
        case "assignStudentsToSheets":
          assignStudentsToSheets(params.spreadsheet_target_url);
          break;
        case "clearAllProfessorSheets":
          clearAllProfessorSheets(params.spreadsheet_target_url);
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
    
    return html.setTitle("Pendaftaran Dosbing Data Controller")
  }
  
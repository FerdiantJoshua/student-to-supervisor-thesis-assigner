const CONST = {
    "COL_NAMES" : {
      "PILIHAN_DOSEN": "Pilihan Dosen",
    },
  
    "SHEET_NAMES": {
      // Assignment sheet
      "FORM_RESPONSES": "Form Responses",
      "TEMPLATE": "Template",
  
      // DB Sheet
      "STUDENTS": "Students",
      "PROFESSORS": "Professors",
      "TOPICS": "Topics",
      "SUPERVISIONS": "Supervisions",
      "SUPERVISION_LEVELS": "SupervisionLevels",
      "OPERATION_PARAMETERS": "OperationParameters",
    },
  
    "MARKERS": {
      "SUPERVISION_LEVEL_PREFIX": "SUPERVISION_LEVEL_",
      "STUDENT_QUEUE": "STUDENT_QUEUE",
    },
    "TEMPLATE_FIRST_DATA_COLUMN": 3, // Start in Column C (Column A is for markers; Column B is for table number)
  };
  
  function test() {
    let ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1TIiC1p0vMptRFfcFISsTG6sJx-W_1Fo8cPD2FJTGv5s/edit")
    let professorSheet = ss.getSheetByName("meow meow");
    return
  }

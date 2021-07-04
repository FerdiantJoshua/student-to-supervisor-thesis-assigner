const CONST = {
  "APP_TIMEOUT_SECONDS": 345,  // Google App Scripts' timeout is 360. This lesser value allows timeout handling for long operation
  "LOCK_WAIT_SECONDS": 3,

  "BATCH_NAME_DEFAULT": "DEFAULT",
  "HTML_DEFAULT_VALUES": {
    "BATCH_NAME": "Batch-A-001",
    "SPREADSHEET_DB_URL": "https://docs.google.com/spreadsheets/d/1bhuvOg6tJOLQyvvyeXDEAUaGwG_dfgavuY6j0fWRjeg/edit",
  },

  "COL_NAMES" : {
    "PILIHAN_DOSEN": "Pilihan Dosen",  // From Registration Form
    "PEMBIMBING_LAIN": "Pembimbing Lain",  // From Template sheet in Assignment spreadsheet
  },

  "SHEET_NAMES": {
    // Assignment sheet
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
  "TEMPLATE_VALID_STUDENTS_COL_GETTER": (templateSheet) => templateSheet.getMaxColumns(),  // Getter for last column

  "CLOSED_SLOT_VALUE": "CLOSED",
  "FIRST_N_VALID_STUDENTS": 40,
};

function test() {
  let ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1TIiC1p0vMptRFfcFISsTG6sJx-W_1Fo8cPD2FJTGv5s/edit")
  let professorSheet = ss.getSheetByName("meow meow");
  return
}

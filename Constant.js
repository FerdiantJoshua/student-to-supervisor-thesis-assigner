const CONST = {
  "BATCH_NAME_ALL": "ALL",

  "COL_NAMES" : {
    // From Registration Form
    "PILIHAN_DOSEN": "Pilihan Dosen",
    // From Template sheet in Assignment spreadsheet
    "PEMBIMBING_LAIN": "Pembimbing Lain",
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
  "FIRST_N_VALID_STUDENTS": 10,
};

function test() {
  let ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1TIiC1p0vMptRFfcFISsTG6sJx-W_1Fo8cPD2FJTGv5s/edit")
  let professorSheet = ss.getSheetByName("meow meow");
  return
}

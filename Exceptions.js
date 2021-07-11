function UnknownOperationException(message) {
  this.message = message;
  this.name = 'UnknownOperationException';
}

function WaitLockTimeoutException(message) {
  this.message = message;
  this.name = 'WaitLockTimeoutException';
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

function PartiallyExecutedOperationException() {}

function isExceptionCustom(err) {
  return err instanceof UnknownOperationException ||
    err instanceof WaitLockTimeoutException ||
    err instanceof BatchNotFoundException ||
    err instanceof AlreadyPreparedSpreadsheetException ||
    err instanceof UnpreparedSpreadsheetException ||
    err instanceof StudentsAlreadyAssignedException ||
    err instanceof ChosenStudentsTableSizeMismatch ||
    err instanceof ChosenStudentsTableDoesNotExist ||
    err instanceof PartiallyExecutedOperationException
}

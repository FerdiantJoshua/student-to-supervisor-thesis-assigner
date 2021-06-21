class Cell {
    constructor(col='A', row=1) {
      this.col = col.toUpperCase()
      this.row = row
    }
  
    getNumberRepr() {
      return [this.col.charCodeAt() - 64, this.row]
    }
  
    getPos() {
      return this.col + this.row
    }
  
    getNextNCols(n) {
      return new Cell(String.fromCharCode(this.col.charCodeAt() + n), this.row)
    }
  
    getNextNRows(n) {
      return new Cell(this.col, this.row + n)
    }
  
    getRangeNextNCols(n) {
      let endCell = this.getNextNCols(n)
      return this.getPos() + ":" + endCell.getPos()
    }
  
    getRangeNextNRows(n) {
      let endCell = this.getNextNRows(n)
      return this.getPos() + ":" + endCell.getPos()
    }
  
    getRangeFullCol() {
      return this.getPos() + ":" + this.col
    }
  
    getRangeFullRow() {
      return this.getPos() + ":" + this.row
    }
  
    getNextEmptyCol(sheet) {
      let fullRow = sheet.getRange(this.getRangeFullRow());
      let values = fullRow.getValues();
      let ct = 0;
      while ( values[0][ct] != "" ) {
        ct++;
      }
      return new Cell(String.fromCharCode(this.col.charCodeAt() + ct), this.row);
    }
  
    getNextEmptyRow(sheet) {
      let fullColumn = sheet.getRange(this.getRangeFullCol());
      let values = fullColumn.getValues();
      let ct = 0;
      while ( values[ct] && values[ct][0] != "" ) {
        ct++;
      }
      return new Cell(this.col, this.row + ct);
    }
  }
  
/**
* @requires getNonEmptyCellCountInRange
* @summary Retrieves the index number of the last row in the given worksheet
* @description A row is considered to be the last row if it is followed by the number of empty rows specified by the threshold argument.
* The index number of the last row is retrieved by manually iterating through each row and finding the last row with data with the given threshold.
* The operation is slow and the function should only be used when necessary.
* @argument {ExcelScript.Worksheet} worksheet Worksheet
* @argument {number} threshold Empty row threshold 
* @returns {number} Returns an integer
*/

function getLastRowIndexInWorksheet(worksheet: ExcelScript.Worksheet, threshold: number) {
  let row = worksheet.getCell(0, 0).getEntireRow();
  let lastRowWithValues: ExcelScript.Range;
  while (threshold > 0) {
    if (getNonEmptyCellCountInRange(row) === 0) {
      threshold = threshold - 1;
    } else {
      lastRowWithValues = row;
    }
    row = row.getRowsBelow();
  }
  return lastRowWithValues.getRowIndex();
}


/**
* @requires getNonEmptyCellCountInRange
* @summary Removes empty rows from the given range
* @argument {ExcelScript.Range} range Range
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function removeEmptyRowsFromRange(range: ExcelScript.Range) {
  let row = range.getLastRow();
  let emptyRows: ExcelScript.Range[] = new Array();
  while (row) {
    if (getNonEmptyCellCountInRange(row) === 0) {
      emptyRows.push(row);
    }
    if (row.getRowIndex() === 0) {
      break;
    }
    row = row.getRowsAbove();
  }
  for (let emptyRow of emptyRows) {
    emptyRow.delete(ExcelScript.DeleteShiftDirection.up);
  }
  return range;
}


/**
* @requires removeEmptyRowsFromRange
* @summary Retrieves a new range starting from the cell with given value 
* @description Throws an error if the range does not contain a cell with the given value.
* @argument {ExcelScript.Range} range Range
* @argument {string} cellValue Cell value
* @argument {boolean} [removeEmptyRows=false] A boolean value indicating whether empty rows should be removed
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getRangeOffsetByCellValue(range: ExcelScript.Range, cellValue: string, removeEmptyRows: boolean = false) {
  const cell = range.find(cellValue, {
    completeMatch: true,
    matchCase: true
  });
  if (typeof cell === 'undefined') {
    throw new Error(`Could not retrieve the cell with value ${cellValue}.`);
  }
  if (removeEmptyRows) {
    return removeEmptyRowsFromRange(range.getOffsetRange(
      cell.getRowIndex(),
      cell.getColumnIndex()
    ).getResizedRange(
      cell.getRowIndex(),
      -cell.getColumnIndex()
    ));
  } else {
    return range.getOffsetRange(
      cell.getRowIndex(),
      cell.getColumnIndex()
    ).getResizedRange(
      cell.getRowIndex(),
      -cell.getColumnIndex()
    );
  }
}


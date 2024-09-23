/**
* @requires removeEmptyRowsFromRange
* @summary Retrieves a new range starting from the given cell address 
* @description Throws an error if the range does not contain a cell with the given address.
* @argument {ExcelScript.Range} range Range
* @argument {string} cellAddress Cell address
* @argument {boolean} [removeEmptyRows=false] A boolean value indicating whether empty rows should be removed
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getRangeOffsetByCellAddress(range: ExcelScript.Range, cellAddress: string, removeEmptyRows: boolean = false) {
  const cell = range.getBoundingRect(cellAddress);
  if (typeof cell === 'undefined') {
    throw new Error(`Could not retrieve the cell with address ${cellAddress}.`);
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

/**
* @requires getRangeOffsetByCellAddress
* @summary Retrieves a new range starting from the cell with given value 
* @description Throws an error if the range does not contain a cell with the given value.
*/

function getRangeOffsetByCellValue(range: ExcelScript.Range, cellValue: string, removeEmptyRows: boolean = false): ExcelScript.Range {
  const cell = range.find(cellValue, {
    completeMatch: true,
    matchCase: true
  });
  if (typeof cell === 'undefined') {
    throw new Error(`Could not retrieve the cell with value ${cellValue}.`);
  }
  return getRangeOffsetByCellAddress(range, cell.getAddress(), removeEmptyRows);
}
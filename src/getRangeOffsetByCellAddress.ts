/**
* @requires removeEmptyRowsFromRange
* @summary Retrieves a new range starting from the given cell address 
* @description Throws an error if the range does not contain a cell with the given address.
*/

function getRangeOffsetByCellAddress(range: ExcelScript.Range, cellAddress: string, removeEmptyRows: boolean = false): ExcelScript.Range {
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
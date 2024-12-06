/**
* @requires normalizeCellValue
* @summary Normalizes an Excel range
*/

function normalizeRangeData(range: ExcelScript.Range, excludeEmptyRows: boolean = true): ExcelRowValue[][] {
  const rangeData: ExcelRowValue[][] = [];
  let rowCount = range.getRowCount() - 1;
  let columnCount = range.getColumnCount();
  for (let rowNumber: number = 0; rowNumber < rowCount; rowNumber++) {
      const row = range.getRow(rowNumber);
      let rowData: ExcelRowValue[] = [];
      for (let columnNumber: number = 0; columnNumber < columnCount; columnNumber++) {
        const cell = row.getColumn(columnNumber);
        rowData.push(normalizeCellValue(cell));
      }
    if (excludeEmptyRows && rowData.every(value => value === null)) {
      continue;
    }
    rangeData.push(rowData);
  }
  return rangeData;
}
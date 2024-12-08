/**
* @summary Sets the given values in a row
*/

function setRowValues(worksheet: ExcelScript.Worksheet, firstRow: number, firstColumn: number, values: (string | number | boolean)[]): ExcelScript.Range {
  const range = worksheet.getRangeByIndexes(firstRow, firstColumn, 1, values.length);
  range.setValues(Array(values));
  return range;
}
/**
* @summary Sets the given values in a column
*/

function setColumnValues(worksheet: ExcelScript.Worksheet, firstRow: number, firstColumn: number, values: (string | number | boolean)[]): ExcelScript.Range {
  const range = worksheet.getRangeByIndexes(firstRow, firstColumn, values.length, 1);
  range.setValues(values.map(value => Array(value)));
  return range;
}
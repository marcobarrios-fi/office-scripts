/**
* @summary Adds a range to the given worksheet as a table 
* @argument {ExcelScript.Range} range Range
* @argument {ExcelScript.Worksheet} worksheet Worksheet
* @argument {boolean} hasHeaders A boolean value indicating whether the created table should have headers
* @returns {ExcelScript.Table} Returns an ExcelScript table
*/

function addRangeAsTableToWorksheet(range: ExcelScript.Range, worksheet: ExcelScript.Worksheet, hasHeaders: boolean) {
  worksheet.getCell(0, 0).copyFrom(range);
  return worksheet.addTable(worksheet.getRangeByIndexes(0, 0, range.getRowCount(), range.getColumnCount()), hasHeaders);
}


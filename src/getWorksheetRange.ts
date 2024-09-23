/**
* @summary Retrieves a new range within the given worksheet 
* @description The new range starts from the first cell and spans the given number of rows and columns.
* @argument {ExcelScript.Worksheet} worksheet Worksheet
* @argument {number} rowCount Row count
* @argument {number} columnCount Column count
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getWorksheetRange(worksheet: ExcelScript.Worksheet, rowCount: number, columnCount: number) {
  return worksheet.getRangeByIndexes(0, 0, rowCount, columnCount);
}


/**
* @requires getWorksheetDataRange
* @summary Retrieves the used data range for the active worksheet 
* @description Returns the first table if the active worksheet contains a table. 
* If the active worksheet does not contain a table, the function returns the used range, which Excel determines automatically. 
* The `ExcelScript.Worksheet.getUsedRange` API contains a bug causing an error when used with XLS files in compatibility mode.
* If there is an error retrieving the used range for the active worksheet, the used range is manually retrieved by iterating through rows and columns.
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getActiveWorksheetDataRange() {
  return getWorksheetDataRange(workbook.getActiveWorksheet().getName());
}
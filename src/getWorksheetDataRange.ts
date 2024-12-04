/**
* @requires getUsedRangeInCompatibilityMode
* @summary Retrieves the used data range for the given worksheet 
* @description Returns the first table if the worksheet contains a table. 
* If the worksheet does not contain a table, the function returns the used range, which Excel determines automatically. 
* The `ExcelScript.Worksheet.getUsedRange` API contains a bug causing an error when used with XLS files in compatibility mode.
* If there is an error retrieving the used range for the active worksheet, the used range is manually retrieved by iterating through rows and columns.
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getWorksheetDataRange(worksheetName: string) {
  // Retrieve the given worksheet
  let worksheet = workbook.getWorksheet(worksheetName);
  // Retrieve the first table in the worksheet
  let firstTable = worksheet.getTables()[0];
  // If the worksheet does not contain a table
  if (typeof firstTable === 'undefined') {
    try {
      // Return the used range for the worksheet
      return worksheet.getUsedRange();
    // If there is an error retrieving the used range
    } catch (error) {
      console.log('Could retrieve the used range.');
      // Retrieve used range manually by iterating through worksheet rows and columns
      return getUsedRangeInCompatibilityMode(worksheet, 10);
    }
  }
  return firstTable.getRange();
}
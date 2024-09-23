/**
* @requires getUsedRangeInCompatibilityMode
* @summary Retrieves the used data range for the active worksheet 
* @description Returns the first table if the active worksheet contains a table. 
* If the active worksheet does not contain a table, the function returns the used range, which Excel determines automatically. 
columns. 
* The `ExcelScript.Worksheet.getUsedRange` API contains a bug causing an error when used with XLS files in compatibility mode.
* If there is an error retrieving the used range for the active worksheet, the used range is manually retrieved by iterating through rows and 
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getActiveWorksheetDataRange() {
  // Retrieve the active sheet
  let activeWorksheet = workbook.getActiveWorksheet();
  // Retrieve the first table in the active worksheet
  let firstTable = activeWorksheet.getTables()[0];
  // If the active worksheet does not contain a table
  if (typeof firstTable === 'undefined') {
    try {
      // Return the used range for the active worksheet
      return activeWorksheet.getUsedRange();
    // If there is an error retrieving the used range
    } catch (error) {
      console.log('Could retrieve the used range.');
      // Retrieve used range manually by iterating through active worksheet rows and columns
      return getUsedRangeInCompatibilityMode(activeWorksheet, 10);
    }
  }
  return firstTable.getRange();
}


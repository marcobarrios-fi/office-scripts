/**
* @requires getLastColumnIndexInWorksheet
* @requires getLastRowIndexInWorksheet
* @requires getWorksheetRange
* @summary Retrieves the used cell range in compatibility mode
* @description The ExcelScript.Worksheet.getUsedRange API contains a bug causing an error when used with XLS files in compatibility mode.
* Microsoft acknowledged the bug in Spetember 27, 2023, according to a forum post.
* The range is retrieved by manually iterating through each cell and finding the last row and column with data with the given threshold.
* The operation is slow and the function should only be used when absolutely necessary.
* @argument {ExcelScript.Worksheet} worksheet Worksheet
* @argument {number} threshold Empty row and column threshold 
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getUsedRangeInCompatibilityMode(worksheet: ExcelScript.Worksheet, threshold: number) {
  return getWorksheetRange(worksheet, getLastRowIndexInWorksheet(worksheet, threshold) + 1, getLastColumnIndexInWorksheet(worksheet, threshold) + 1);
}


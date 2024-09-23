/**
* @requires getNonEmptyCellCountInRange
* @summary Retrieves the index number of the last column in the given worksheet
* @description A column is considered to be the last column if it is followed by the number of empty columns specified by the threshold argument.
* The index number of the last column is retrieved by manually iterating through each column and finding the last column with data with the given threshold.
* The operation is slow and the function should only be used when necessary.
* @argument {ExcelScript.Worksheet} worksheet Worksheet
* @argument {number} threshold Empty column threshold 
* @returns {number} Returns an integer
*/
  
function getLastColumnIndexInWorksheet(worksheet: ExcelScript.Worksheet, threshold: number) {
  let column = worksheet.getCell(0, 0).getEntireColumn();
  let lastColumnWithValues: ExcelScript.Range;
  while (threshold > 0) {
    if (getNonEmptyCellCountInRange(column) === 0) {
      threshold = threshold - 1;
    } else {
      lastColumnWithValues = column;
    }
    column = column.getColumnsAfter();
  }
  return lastColumnWithValues.getColumnIndex();
}


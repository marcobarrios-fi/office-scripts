/**
* @summary Adds a new worksheet to the workbook
* @description If a worksheet with the given name already exists, it will be replaced with an empty worksheet.
* The name of the worksheet is automatically truncated to the maximum length supported by Excel.
* The maximum length for worksheet names in Excel is 31 characters.
* @argument {string} worksheetName Worksheet name
* @returns {ExcelScript.Worksheet} Returns an ExcelScript worksheet
*/

function addWorksheet(worksheetName: string) {
  worksheetName = worksheetName.slice(0, 30);
  if (workbook.getWorksheet(worksheetName)) {
    workbook.getWorksheet(worksheetName).delete();
  }
  return workbook.addWorksheet(worksheetName);
}


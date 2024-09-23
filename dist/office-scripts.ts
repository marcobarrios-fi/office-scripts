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

/**
* @summary Adds a new worksheet to the workbook
* @description If a worksheet with the given name already exists, it will be deleted.
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

/**
* @summary Retrieves values of the first column in the given range
* @argument {ExcelScript.Range} range Range
* @returns {array} Returns an array
*/

function getFirstColumnValuesInRange(range: ExcelScript.Range) {
  return range.getColumn(0).getTexts().map(columnValue => columnValue[0]);
}

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

/**
* @requires getNonEmptyCellCountInRange
* @summary Retrieves the index number of the last row in the given worksheet
* @description A row is considered to be the last row if it is followed by the number of empty rows specified by the threshold argument.
* The index number of the last row is retrieved by manually iterating through each row and finding the last row with data with the given threshold.
* The operation is slow and the function should only be used when necessary.
* @argument {ExcelScript.Worksheet} worksheet Worksheet
* @argument {number} threshold Empty row threshold 
* @returns {number} Returns an integer
*/

function getLastRowIndexInWorksheet(worksheet: ExcelScript.Worksheet, threshold: number) {
  let row = worksheet.getCell(0, 0).getEntireRow();
  let lastRowWithValues: ExcelScript.Range;
  while (threshold > 0) {
    if (getNonEmptyCellCountInRange(row) === 0) {
      threshold = threshold - 1;
    } else {
      lastRowWithValues = row;
    }
    row = row.getRowsBelow();
  }
  return lastRowWithValues.getRowIndex();
}

/**
* @summary Retrieves the number of non-emptpy cells within the given range
* @argument {ExcelScript.Range} range Range
* @returns {number} Returns an integer
*/

 function getNonEmptyCellCountInRange(range: ExcelScript.Range) {
  return range.getTexts().filter(cellValue => cellValue[0].trim().length).length;
}

/**
* @requires removeEmptyRowsFromRange
* @summary Retrieves a new range starting from the given cell address 
* @description Throws an error if the range does not contain a cell with the given address.
* @argument {ExcelScript.Range} range Range
* @argument {string} cellAddress Cell address
* @argument {boolean} [removeEmptyRows=false] A boolean value indicating whether empty rows should be removed
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getRangeOffsetByCellAddress(range: ExcelScript.Range, cellAddress: string, removeEmptyRows: boolean = false) {
  const cell = range.getBoundingRect(cellAddress);
  if (typeof cell === 'undefined') {
    throw new Error(`Could not retrieve the cell with address ${cellAddress}.`);
  }
  if (removeEmptyRows) {
    return removeEmptyRowsFromRange(range.getOffsetRange(
      cell.getRowIndex(),
      cell.getColumnIndex()
    ).getResizedRange(
      cell.getRowIndex(),
      -cell.getColumnIndex()
    ));
  } else {
    return range.getOffsetRange(
      cell.getRowIndex(),
      cell.getColumnIndex()
    ).getResizedRange(
      cell.getRowIndex(),
      -cell.getColumnIndex()
    );
  }
}
/**
* @requires removeEmptyRowsFromRange
* @summary Retrieves a new range starting from the cell with given value 
* @description Throws an error if the range does not contain a cell with the given value.
* @argument {ExcelScript.Range} range Range
* @argument {string} cellValue Cell value
* @argument {boolean} [removeEmptyRows=false] A boolean value indicating whether empty rows should be removed
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function getRangeOffsetByCellValue(range: ExcelScript.Range, cellValue: string, removeEmptyRows: boolean = false) {
  const cell = range.find(cellValue, {
    completeMatch: true,
    matchCase: true
  });
  if (typeof cell === 'undefined') {
    throw new Error(`Could not retrieve the cell with value ${cellValue}.`);
  }
  if (removeEmptyRows) {
    return removeEmptyRowsFromRange(range.getOffsetRange(
      cell.getRowIndex(),
      cell.getColumnIndex()
    ).getResizedRange(
      cell.getRowIndex(),
      -cell.getColumnIndex()
    ));
  } else {
    return range.getOffsetRange(
      cell.getRowIndex(),
      cell.getColumnIndex()
    ).getResizedRange(
      cell.getRowIndex(),
      -cell.getColumnIndex()
    );
  }
}

/**
* @summary Retrieves column values in the given table
* @argument {ExcelScript.Table} table Table
* @argument {string} columnName Column name
* @returns {Array.<string>} Returns an array of strings
*/

function getTableColumnValuesByColumnName(table: ExcelScript.Table, columnName: string) {
  const column = table.getColumnByName(columnName);
  if (typeof column === 'undefined') {
    throw new Error(`Could not retrieve ${columnName} column.`);
  }
  return column.getRangeBetweenHeaderAndTotal().getTexts().map(columnValue => columnValue[0]);
}

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

/**
* @requires getNonEmptyCellCountInRange
* @summary Removes empty rows from the given range
* @argument {ExcelScript.Range} range Range
* @returns {ExcelScript.Range} Returns an ExcelScript range
*/

function removeEmptyRowsFromRange(range: ExcelScript.Range) {
  let row = range.getLastRow();
  let emptyRows: ExcelScript.Range[] = new Array();
  while (row) {
    if (getNonEmptyCellCountInRange(row) === 0) {
      emptyRows.push(row);
    }
    if (row.getRowIndex() === 0) {
      break;
    }
    row = row.getRowsAbove();
  }
  for (let emptyRow of emptyRows) {
    emptyRow.delete(ExcelScript.DeleteShiftDirection.up);
  }
  return range;
}


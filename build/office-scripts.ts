/* Helper Functions (Build 2024-12-08) */

/**
* @summary ExcelDate class
*/

class ExcelDate extends Date {

  /**
  * @summary Constructs a JavaScript Date object from an Excel date value
  * @argument {number} excelDateValue Excel date value
  * @see https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/javascript-dates
  */

  constructor(excelDateValue: number) { 
    super(Math.round((excelDateValue - 25569) * 86400 * 1000));
  }

  /**
  * @summary Returns the date in the ISO 8601 format
  * @example 2024-12-01T12:30:00.000Z
  */
  
  toString(): string {
    return super.toISOString();
  }

  /**
  * @summary Returns the date in the US long format
  * @example December 1, 2024
  */
  
  toUSLongFormatString(): string {
    return super.toLocaleDateString('en-US', {
      month: 'long', day: 'numeric', year: 'numeric',
    });
  }

}

/**
* @summary ExcelURL class
* The ExcelURL class extends the standard JavaScript URL object.
*/

class ExcelURL extends URL {

  /**
  * @summary URL title
  */

  title: string;

  /**
  * @summary Constructs a custom JavaScript URL object from an Excel hyperlink
  */

  constructor(hyperlink: ExcelScript.RangeHyperlink) {
    super(hyperlink.address);
    this.title = hyperlink.textToDisplay;
  }

  /**
  * @summary Constructs an Excel hyperlink
  */

  toExcelHyperlink(): ExcelScript.RangeHyperlink {
    return {
      address: this.href,
      screenTip: this.title,
      textToDisplay: this.title
    }
  }

}

/**
* ExcelValue type definition
*/

type ExcelValue = boolean | null | number | string | ExcelDate | ExcelURL;

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

/**
* @summary Capitalizes the given name
* @example Example of capitalizing a name:
* capitalizeName('marco barrios');
* > Marco Barrios
*/
 
function capitalizeName(name: string): string {
  return name.toLowerCase().replace(new RegExp('[A-Za-zÀ-ÖØ-öø-ÿ]+', 'g'), function (substring) {
    if (Array('de', 'der', 'dit', 'van', 'von').includes(substring)) {
      return substring;
    } else if (substring.startsWith('mc') && substring.length > 3) {
      return 'Mc' + substring[2].toUpperCase() + substring.slice(3);
    } else if (substring.startsWith('mac') && substring.length > 4) {
      return 'Mac' + substring[3].toUpperCase() + substring.slice(4);
    }
    return substring[0].toUpperCase() + substring.slice(1);
  });
}

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

/**
* @summary Retrieves column values in a data range
* @description Throws an error if a column with a given name does not exist.
*/

function getDataRangeColumn(dataRange: ExcelValue[][], columnName: string) {
  if (!dataRange[0].includes(columnName)) {
    throw new Error(`Column ${columnName} does not exist`);
  }
  return dataRange.slice(1).map(row => row[dataRange[0].indexOf(columnName)]);
}

/**
* @requires getDataRangeColumn
* @summary Retrieves unique column values in a data range
* @description Throws an error if a column with a given name does not exist.
*/

function getDataRangeColumnUniqueValues(dataRange: ExcelValue[][], columnName: string) {
  return getDataRangeColumn(dataRange, columnName).filter((value, index, array) => array.indexOf(value) === index);
}

/**
* @summary Retrieves all columns in a data range
* @description The range is returned as a JavaScript Map object
*/

function getDataRangeColumns(dataRange: ExcelValue[][]) {
  return new Map(dataRange[0].map((columnName, columnNumber) => 
    [columnName as string, dataRange.slice(1).map(row => 
      row[columnNumber])
    ])
  );
}

/**
* @summary Retrieves all rows in a data range
* @description The range is returned as an array consisting of JavaScript Map objects
*/

function getDataRangeRows(dataRange: ExcelValue[][]) {
  return dataRange.slice(1).map(row => 
    new Map(dataRange[0].map((columnName, columnNumber) => 
      [columnName, row[columnNumber]]
    ))
  );
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
  
function getLastColumnIndexInWorksheet(worksheet: ExcelScript.Worksheet, threshold: number): number {
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
* @summary Retrieves a new range starting from the cell with given value 
* @description Throws an error if the range does not contain a cell with the given value.
*/

function getRangeOffsetByCellValue(range: ExcelScript.Range, cellValue: string, removeEmptyRows: boolean = false): ExcelScript.Range {
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
      console.log(`Could retrieve the used range.`);
      // Retrieve used range manually by iterating through worksheet rows and columns
      return getUsedRangeInCompatibilityMode(worksheet, 10);
    }
  }
  return firstTable.getRange();
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
* @summary Parses the given `ExcelScript.Range` range to an array of arrays.
* @description Empty cells are converted to `null` values, dates are converted to `ExcelDate` objects, and hyperlinks are optionally converted to `ExcelURL` objects.
* Please note that reading data from the Excel using the Office Scripts API is extremely slow. 
* Retrieving value types and number formats by row would make the script run significantly slower.
* The Office Scripts API does not provide any method for retrieving or identifying all hyperlinks so they must be retrieved one cell at the time. 
* Note: Enabling parsing hyperlinks causes the script to run over hundred times slower.
*/

function parseRange(range: ExcelScript.Range, 
  excludeEmptyRows: boolean = true, parseHyperlinks: boolean = false): ExcelValue[][] {
  const rangeProperties = {
   valueTypes: range.getValueTypes(),
   numberFormats: range.getNumberFormatCategories()
  }
  return range.getValues().map((row, rowNumber) => {
   if (excludeEmptyRows) {
    if (rangeProperties.valueTypes[rowNumber].every(
      valueType => valueType === ExcelScript.RangeValueType.empty
    )) {
     return null;
    }
   }
   return row.map((column, columnNumber) => {
    switch (rangeProperties.valueTypes[rowNumber][columnNumber]) {
     case ExcelScript.RangeValueType.empty:
      return null;
     case ExcelScript.RangeValueType.string:
      if (parseHyperlinks) {
       const hyperlink = range.getCell(rowNumber, columnNumber).getHyperlink();
       if (hyperlink) {
        return new ExcelURL(hyperlink);
       }
      }
      if ((column as string).trim().length) {
       return (column as string).trim();
      } else {
       return null;
      }
     case ExcelScript.RangeValueType.double:
      switch (rangeProperties.numberFormats[rowNumber][columnNumber as number]) {
       case ExcelScript.NumberFormatCategory.date:
        return new ExcelDate(column as number);
      }
      return column;
    }
    return column;
   })
  }).filter(row => Boolean(row));
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

/**
* @summary Sets the given values in a column
*/

function setColumnValues(worksheet: ExcelScript.Worksheet, firstRow: number, firstColumn: number, values: (string | number | boolean)[]): ExcelScript.Range {
  const range = worksheet.getRangeByIndexes(firstRow, firstColumn, values.length, 1);
  range.setValues(values.map(value => Array(value)));
  return range;
}

/**
* @summary Sets range border with the given color and weight
* @details Border weight must be `hairline`, `medium`, `thick`, or `thin`.
* @argument {string} [borderColor='#000000'] Border color in hexadecimal value (default color is black) 
* @argument {string} [borderWeight='medium'] Border weight (default weight is medium)
* @returns {void} Returns undefined
*/

function setRangeBorder(range: ExcelScript.Range, borderColor: string = '#000000', borderWeight: keyof typeof ExcelScript.BorderWeight = 'medium') {
  const topBorder = range.getRow(0).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop);
  const bottomBorder = range.getRow(range.getRowCount() - 1).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
  const leftBorder = range.getColumn(0).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft);
  const rightBorder = range.getColumn(range.getColumnCount() - 1).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight);
  // Top border color and weight
  topBorder.setColor(borderColor);
  topBorder.setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Bottom border color and weight
  bottomBorder.setColor(borderColor);
  bottomBorder.setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Left border color and weight
  leftBorder.setColor(borderColor);
  leftBorder.setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Right border color and weight
  rightBorder.setColor(borderColor);
  rightBorder.setWeight(ExcelScript.BorderWeight[borderWeight]);
}

/**
* @summary Sets the given values in a row
*/

function setRowValues(worksheet: ExcelScript.Worksheet, firstRow: number, firstColumn: number, values: (string | number | boolean)[]): ExcelScript.Range {
  const range = worksheet.getRangeByIndexes(firstRow, firstColumn, 1, values.length);
  range.setValues(Array(values));
  return range;
}


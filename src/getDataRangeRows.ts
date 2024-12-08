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
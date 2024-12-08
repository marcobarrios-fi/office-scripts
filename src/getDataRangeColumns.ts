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
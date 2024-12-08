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
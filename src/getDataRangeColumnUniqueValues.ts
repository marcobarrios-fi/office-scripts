/**
* @requires getDataRangeColumn
* @summary Retrieves unique column values in a data range
* @description Throws an error if a column with a given name does not exist.
*/

function getDataRangeColumnUniqueValues(dataRange: ExcelValue[][], columnName: string) {
  return getDataRangeColumn(dataRange, columnName).filter((value, index, array) => array.indexOf(value) === index);
}
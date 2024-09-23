/**
* @summary Retrieves values of the first column in the given range
* @argument {ExcelScript.Range} range Range
* @returns {array} Returns an array
*/

function getFirstColumnValuesInRange(range: ExcelScript.Range) {
  return range.getColumn(0).getTexts().map(columnValue => columnValue[0]);
}


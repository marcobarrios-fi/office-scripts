/**
* @summary Retrieves the number of non-emptpy cells within the given range
* @argument {ExcelScript.Range} range Range
* @returns {number} Returns an integer
*/

 function getNonEmptyCellCountInRange(range: ExcelScript.Range) {
  return range.getTexts().filter(cellValue => cellValue[0].trim().length).length;
}


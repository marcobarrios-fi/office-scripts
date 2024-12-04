/**
* @requires getObjectFromKeysAndValues
* @summary Converts the given Excel range to an array of objects
* @details The values of the first row are used as the the object keys.
* Office Scripts does not support JavaScript `Object.fromEntries` function.
* @argument {ExcelScript.Range} range Excel range
* @returns {array} Returns an array of objects
*/

function getRangeAsArrayOfObjects(range: ExcelScript.Range) {
  return range.getOffsetRange(1, 0).getValues().map(values =>
    getObjectFromKeysAndValues(range.getRow(0).getTexts()[0], values)
  );
}
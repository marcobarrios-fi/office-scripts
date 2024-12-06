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
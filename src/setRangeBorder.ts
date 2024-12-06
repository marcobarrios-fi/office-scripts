/**
* @summary Sets range border with the given weight and color
* @details Border weight must be `hairline`, `medium`, `thick`, or `thin`.
* @argument {string} [borderColor='#000000'] Border color in hexadecimal value (default color is black) 
* @argument {string} [borderWeight='medium'] Border weight (default weight is medium)
* @returns {void} Returns undefined
*/

function setRangeBorder(
  range: ExcelScript.Range, 
  borderColor: string = '#000000', 
  borderWeight: keyof typeof ExcelScript.BorderWeight = 'medium'
) {
  // Top border weight
  range.getRow(0).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeTop).
    setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Bottom border weight
  range.getRow(range.getRowCount() - 1).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeBottom).
    setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Left border weight
  range.getColumn(0).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeLeft).
    setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Right border weight
  range.getColumn(range.getColumnCount() - 1).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeRight).
    setWeight(ExcelScript.BorderWeight[borderWeight]);
  // Top border color
  range.getRow(0).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeTop).
    setColor(borderColor);
  // Bottom border color
  range.getRow(range.getRowCount() - 1).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeBottom).
    setColor(borderColor);
  // Left border color
  range.getColumn(0).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeLeft).
    setColor(borderColor);
  // Right border color
  range.getColumn(range.getColumnCount() - 1).getFormat().
    getRangeBorder(ExcelScript.BorderIndex.edgeRight).
    setColor(borderColor);
}
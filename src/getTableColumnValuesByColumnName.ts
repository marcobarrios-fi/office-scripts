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


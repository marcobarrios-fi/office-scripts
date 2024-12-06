/**
* @summary ExcelRange class
*/

class ExcelRange {

  rows: ExcelRow[];

  /**
  * @summary Constructs a new `ExcelRange` object
  */

  constructor(range: ExcelScript.Range) {
    this.rows = this.parseExcelRange(range);
    return this;
  }

  /**
  * @requires ExcelDate
  * @requires ExcelURL
  * @summary Parses an `ExcelScript.Range` cell to an `ExcelRowValue` value
  */

  parseExcelCell(cell: ExcelScript.Range): ExcelRowValue {
    switch (cell.getValueType()) {
      case ExcelScript.RangeValueType.boolean:
        if (cell.getText() === 'TRUE') {
          return true;
        } else {
          return false;
        }
      case ExcelScript.RangeValueType.double:
        if (cell.getNumberFormatCategory() === ExcelScript.NumberFormatCategory.date) {
          return new ExcelDate(cell.getValue() as number);
        } else {
          return cell.getValue();
        }
      case ExcelScript.RangeValueType.empty:
        return null;
      case ExcelScript.RangeValueType.string:
        const cellHyperlink = cell.getHyperlink();
        if (cellHyperlink) {
          return new ExcelURL(cellHyperlink);
        } else {
          const cellValue = cell.getText();
          if (cellValue.trim().length) {
            return cellValue.trim();
          } else {
            return null;
          }
        }
      default: 
        console.log(`Unsupported data type in cell ${cell.getAddress()}`);
        return cell.getText();
    }
  }

  /**
  * @summary Parses an `ExcelScript.Range` row to an array consisting of `ExcelRowValue` values
  */

  parseExcelRow(row: ExcelScript.Range): ExcelRowValue[] {
    return Array.from(Array(row.getColumnCount())).map((_, columnNumber) => 
      this.parseExcelCell(row.getColumn(columnNumber))
    );
  }

  /**
  * @requires ExcelRow
  * @summary Parses an `ExcelScript.Range` range to an array consisting of `ExcelRow` objects
  */

  parseExcelRange(range: ExcelScript.Range): ExcelRow {
    const keys = range.getRow(0).getTexts()[0];
    range = range.getOffsetRange(1, 0);
    return Array.from(Array(range.getRowCount())).map((_, rowNumber) => 
      new ExcelRow(keys, this.parseExcelRow(range.getRow(rowNumber)))
    ).filter(row => Object.values(row).every(value => value === null) === false)
  }
  
}
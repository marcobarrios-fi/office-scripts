/**
* @summary Normalizes an Excel cell value
*/

function normalizeCellValue(cell: ExcelScript.Range): ExcelRowValue {
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
      if (cell.getHyperlink()) {
        return new ExcelURL(cellHyperlink.address, cellHyperlink.textToDisplay);
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
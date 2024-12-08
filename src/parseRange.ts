/**
* @summary Parses the given `ExcelScript.Range` range to an array of arrays.
* @description Empty cells are converted to `null` values, dates are converted to `ExcelDate` objects, and hyperlinks are optionally converted to `ExcelURL` objects.
* Please note that reading data from the Excel using the Office Scripts API is extremely slow. 
* Retrieving value types and number formats by row would make the script run significantly slower.
* The Office Scripts API does not provide any method for retrieving or identifying all hyperlinks so they must be retrieved one cell at the time. 
* Note: Enabling parsing hyperlinks causes the script to run over hundred times slower.
*/

function parseRange(range: ExcelScript.Range, 
  excludeEmptyRows: boolean = true, parseHyperlinks: boolean = false): ExcelValue[][] {
  const rangeProperties = {
   valueTypes: range.getValueTypes(),
   numberFormats: range.getNumberFormatCategories()
  }
  return range.getValues().map((row, rowNumber) => {
   if (excludeEmptyRows) {
    if (rangeProperties.valueTypes[rowNumber].every(
      valueType => valueType === ExcelScript.RangeValueType.empty
    )) {
     return null;
    }
   }
   return row.map((column, columnNumber) => {
    switch (rangeProperties.valueTypes[rowNumber][columnNumber]) {
     case ExcelScript.RangeValueType.empty:
      return null;
     case ExcelScript.RangeValueType.string:
      if (parseHyperlinks) {
       const hyperlink = range.getCell(rowNumber, columnNumber).getHyperlink();
       if (hyperlink) {
        return new ExcelURL(hyperlink);
       }
      }
      if ((column as string).trim().length) {
       return (column as string).trim();
      } else {
       return null;
      }
     case ExcelScript.RangeValueType.double:
      switch (rangeProperties.numberFormats[rowNumber][columnNumber as number]) {
       case ExcelScript.NumberFormatCategory.date:
        return new ExcelDate(column as number);
      }
      return column;
    }
    return column;
   })
  }).filter(row => Boolean(row));
}
# Automating Excel using Office Scripts

## Why Office Scripts?
[Office Scripts in Excel](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) allows automating day-to-day tasks in Excel using [TypeScript](https://www.typescriptlang.org) scripting language. TypeScript is a superset of JavaScript and supports static typing. 

Some of the functionality provided by Office Scripts can be achieved by using Excel formulas, but it often requires re-applying the formulas to Excel files authored by third-party software or other users.

Office Scripts are available for Microsoft 365 commercial and education licenses with access to desktop applications, and can be automated using the [Microsoft Power Automate](https://www.microsoft.com/en-us/power-platform/products/power-automate) platform.

## Files

This repository contains helper functions for writing Office Scripts in Excel.

The [`dist/office-scripts.ts`](../dist/office-scripts.ts) file contains all the functions in the repository. 

The [`dist/office-scripts.min.ts`](../dist/office-scripts.min.ts) file contains all the functions in a minified format.

The minified file was generated with the [Novel Tools TypeScript Minifier](https://www.novel.tools/minify/typescript/).

## Getting Started

1. Open Excel, select the **Automate** tab in the Ribbon, click the **New Script** button, and click the **Write a Script** button in the Script Editor.

2. Copy the contents of the [`dist/office-scripts.min.ts`](../dist/office-scripts.min.ts) file to inside the **main** function.

## Functions

### [`getActiveWorksheetDataRange()`](../src/getActiveWorksheetDataRange.ts)

Helper function for retrieving the used data range for the active worksheet.

Returns the first table if the active worksheet contains a table. If the active worksheet does not contain a table, the function returns the used range, which Excel determines automatically. The `ExcelScript.`&ZeroWidthSpace;`Worksheet.`&ZeroWidthSpace;`getUsedRange` API contains a bug causing an error when used with XLS files in compatibility mode. If there is an error retrieving the used range for the active worksheet, the used range is manually retrieved by iterating through rows and columns.

> Returns an ExcelScript range (`ExcelScript.Range`)

### [`getRangeOffsetByCellAddress(range, cellAddress, [removeEmptyRows])`](../src/getRangeOffsetByCellAddress.ts)

Retrieves a new range starting from the given cell address.

> Returns an ExcelScript range (`ExcelScript.Range`)

### [`getRangeOffsetByCellValue(range, cellValue, [removeEmptyRows])`](../src/getRangeOffsetByCellValue.ts)

Retrieves a new range starting from the cell with the given value.

> Returns an ExcelScript range (`ExcelScript.Range`)

### [`getFirstColumnValuesInRange(range)`](../src/getFirstColumnValuesInRange.ts)

Retrieves values of the first column in the given range.

> Returns an array of strings (`Array[string]`)

### [`addWorksheet(worksheetName)`](../src/addWorksheet.ts)

Adds a new worksheet to the workbook.

If a worksheet with the given name already exists, it will be replaced with an empty worksheet. The name of the worksheet is automatically truncated to the maximum length supported by Excel. The maximum length for worksheet names in Excel is 31 characters.

> Returns an ExcelScript worksheet (`ExcelScript.Worksheet`)

### [`addRangeAsTableToWorksheet(range, worksheet, hasHeaders)`](../src/addRangeAsTableToWorksheet.ts)

Adds a range to the given worksheet as a table.

> Returns an ExcelScript table (`ExcelScript.Table`)

### [`getTableColumnValuesByColumnName(table, columnName)`](../src/getTableColumnValuesByColumnName.ts)

Retrieves column values in the given table.

> Returns an array of strings (`Array[string]`)

## Tutorial

### Example Data

Below is an example Excel worksheet containing company quarterly results. 

<table>
  <tr>
    <th align="left" width="211">A</th>
    <th align="left" width="211">B</th>
    <th align="left" width="211">C</th>
    <th align="left" width="211">D</th>
  </tr>
  <tr>
    <td>Quarterly Results</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Q1/2024</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
   <tr>
    <td><strong>Region</strong></td>
    <td><strong>Country</strong></td>
    <td><strong>Revenue</strong></td>
    <td><strong>Profit</strong></td>
  </tr>
  <tr>
    <td>EMEA</td>
    <td>Andorra</td>
    <td>100,000</td>
    <td>10,000</td>
  </tr>
   <tr>
    <td>EMEA</td>
    <td>United Arab Emirates</td>
    <td>200,000</td>
    <td>20,000</td>
  </tr>
  <tr>
    <td>EMEA</td>
    <td>Afghanistan</td>
    <td>300,000</td>
    <td>30,000</td>
  </tr>
   <tr>
    <td>LATAM</td>
    <td>Antigua and Barbuda</td>
    <td>400,000</td>
    <td>40,000</td>
  </tr>
</table>

### Retrieving Data Range

To process data with Office Scripts in Excel, we must determine which cells in the worksheet contain data. We can use the [`getActiveWorksheetDataRange`](../src/getActiveWorksheetDataRange.ts) function to retrieve the used range for the currently active worksheet:

```
function main(workbook: ExcelScript.Workbook) {

  let usedRange = getActiveWorksheetDataRange();

}
```

Excel worksheets often contain metadata and other information that we can disregard. To retrieve the relevant data range, we can use the [`getRangeOffsetByCellValue`](../src/getRangeOffsetByCellValue.ts) function:

```
function main(workbook: ExcelScript.Workbook) {
  
  let usedRange = getActiveWorksheetDataRange();
  let dataRange = getRangeOffsetByCellValue(usedRange, 'Region');

}
```

Alternatively, we can use the [`getRangeOffsetByCellAddress`](../src/getRangeOffsetByCellAddress.ts) function:

```
function main(workbook: ExcelScript.Workbook) {
  
  let usedRange = getActiveWorksheetDataRange();
  let dataRange = getRangeOffsetByCellAddress(usedRange, 'A3');

}
```

Now we have an Excel range containing only the relevant data:

<table>
  <tr>
    <td width="211"><strong>Region</strong></td>
    <td width="211"><strong>Country</strong></td>
    <td width="211"><strong>Revenue</strong></td>
    <td width="211"><strong>Profit</strong></td>
  </tr>
  <tr>
    <td>EMEA</td>
    <td>Andorra</td>
    <td>100,000</td>
    <td>10,000</td>
  </tr>
  <tr>
    <td>EMEA</td>
    <td>United Arab Emirates</td>
    <td>200,000</td>
    <td>20,000</td>
  </tr>
  <tr>
    <td>EMEA</td>
    <td>Afghanistan</td>
    <td>300,000</td>
    <td>30,000</td>
  </tr>
  <tr>
    <td>LATAM</td>
    <td>Antigua and Barbuda</td>
    <td>400,000</td>
    <td>40,000</td>
  </tr>
</table>

### Processing Data

To retrieve values of the first column, we can use the [`getFirstColumnValuesInRange`](../src/getFirstColumnValuesInRange.ts) function:

```
function main(workbook: ExcelScript.Workbook) {
  
  let usedRange = getActiveWorksheetDataRange();
  let dataRange = getRangeOffsetByCellValue(usedRange, 'Region');
  
  let regions = getFirstColumnValuesInRange(dataRange);
  
  regions = regions.filter((region, index, regions) => 
    regions.indexOf(region) === index
  );

  console.log(regions);

}
```

> Output: `['EMEA', 'LATAM']` Array(2)

To add the data to a new worksheet as a table, we can use the [`addWorksheet`](../src/addWorksheet.ts) and [`addRangeAsTableToWorksheet`](../src/addRangeAsTableToWorksheet.ts) functions:

```
function main(workbook: ExcelScript.Workbook) {
  
  let usedRange = getActiveWorksheetDataRange();
  let dataRange = getRangeOffsetByCellValue(usedRange, 'Region');
  
  let dataWorksheet = addWorksheet('My Data');
  let quarterlyResults = addRangeAsTableToWorksheet(dataRange, dataWorksheet, true);

}
```

Processing data in Excel is easier when the data is contained in tables. If we want to retrieve values in a column, we can use the [`getTableColumnValuesByColumnName`](../src/getTableColumnValuesByColumnName.ts) function:

```
function main(workbook: ExcelScript.Workbook) {
  
  let usedRange = getActiveWorksheetDataRange();
  let dataRange = getRangeOffsetByCellValue(usedRange, 'Region');
  
  let dataWorksheet = addWorksheet('My Data');
  let quarterlyResults = addRangeAsTableToWorksheet(dataRange, dataWorksheet, true);

  let profits = getTableColumnValuesByColumnName(quarterlyResults, 'Profits');

  console.log(profits);

}
```

> Output: `['10,000', '20,000', '30,000', '40,000']` Array(4)


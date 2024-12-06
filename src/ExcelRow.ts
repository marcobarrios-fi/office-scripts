/**
* @summary ExcelRow class
*/

class ExcelRow {

  /**
  * @summary Constructs a new `ExcelRow` object using the specified keys and values
  * @details Values must match the `ExcelRowValue` type definition
  * Office Scripts does not support JavaScript `Object.fromEntries` function.
  */

  constructor(keys: string[], values: ExcelRowValue[]) {
    for (const index in keys) {
      if (keys[index].trim().length) {
        this[keys[index].trim()] = values[index]
      }
    }
  }
  
}
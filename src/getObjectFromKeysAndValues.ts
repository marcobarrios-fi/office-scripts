/**
* @summary Creates an object using the specified keys and values
* @details Values must match the `ExcelRowValue` type definition
* Office Scripts does not support JavaScript `Object.fromEntries` function.
*/

function getObjectFromKeysAndValues(keys: string[], values: ExcelRowValue[]): {} {
  let object = new Object();
  for (const index in keys) {
    if (keys[index].trim().length) {
      object[keys[index].trim()] = values[index]
    }
  }
  return object;
}
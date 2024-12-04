/**
* @summary Creates an object using the specified keys and values
* @details Values must be either strings, numbers, or boolean values.
* Office Scripts does not support JavaScript `Object.fromEntries` function.
* @argument {array} keys Keys
* @argument {array} values Values
* @returns {object} Returns an objects
*/

function getObjectFromKeysAndValues(keys: string[], values: (string|number|boolean)[]) {
  let object = new Object();
    for (const index in keys) {
      if (keys[index].trim().length) {
        object[keys[index].trim()] = values[index];
      }
    }
  return object;
}
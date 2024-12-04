/**
* @summary Retrieves unique array values
* @argument {array} array Array
* @returns {array} Returns an array
*/

function getUniqueArrayValues(array: []) {
  return array.filter((
    value, index, array
  ) => array.indexOf(value) === index);
}
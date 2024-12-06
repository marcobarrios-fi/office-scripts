/**
* @summary Capitalizes the given name
* @example Example of capitalizing a name:
* capitalizeName('marco barrios');
* > Marco Barrios
*/
 
function capitalizeName(name: string): string {
  return name.toLowerCase().replace(new RegExp('[A-Za-zÀ-ÖØ-öø-ÿ]+', 'g'), function (substring) {
    if (Array('de', 'der', 'dit', 'van', 'von').includes(substring)) {
      return substring;
    } else if (substring.startsWith('mc') && substring.length > 3) {
      return 'Mc' + substring[2].toUpperCase() + substring.slice(3);
    } else if (substring.startsWith('mac') && substring.length > 4) {
      return 'Mac' + substring[3].toUpperCase() + substring.slice(4);
    }
    return substring[0].toUpperCase() + substring.slice(1);
  });
}
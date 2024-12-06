/**
* @summary ExcelDate class
*/

class ExcelDate extends Date {

  /**
  * @summary Constructs a JavaScript Date object from an Excel date value
  * @argument {number} excelDateValue Excel date value
  * @see https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/javascript-dates
  */

  constructor(excelDateValue: number) { 
    super(Math.round((excelDateValue - 25569) * 86400 * 1000));
  }

  /**
  * @summary Returns the date in the ISO 8601 format
  * @example 2024-12-01T12:30:00.000Z
  */
  
  toString(): string {
    return this.toISOString();
  }

  /**
  * @summary Returns the date in the US long format
  * @example December 1, 2024
  */
  
  toUSLongFormatString(): string {
    return this.toLocaleDateString('en-US', {
      month: 'long', day: 'numeric', year: 'numeric',
    });
  }

}
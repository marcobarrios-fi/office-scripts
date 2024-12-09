/**
* @summary ExcelDate class
* @see https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/javascript-dates
* @see https://stackoverflow.com/questions/46130132/converting-unix-time-into-date-time-via-excel
*/

class ExcelDate extends Date {

  /**
  * @summary Time format difference between JavaScript and Excel date formats.
  * @description JavaScript uses Unix epoch for calculating dates. Unix epoch is the number of non-leap seconds passed since January 1, 1970.
  * Instead of using Unix epoch, Excel uses January 1, 1900, for calculating dates. The difference between these two dates is 25 569 days.
  */

  timeFormatDifference = 25569;

  /**
  * @summary Helper method for converting days to millisecons
  * @description  Multiplies days with hours, minutes, seconds, and milliseconds.
  */

  daysToMilliseconds(days: number) {
    return days * (24 * 60 * 60 * 1000);
  }

  /**
  * @summary Helper method for converting milliseconds to days
  * @description Divides milliseconds by hours, minutes, seconds, and milliseconds.
  */

  millisecondsToDays(milliseconds: number) {
    return milliseconds / (24 * 60 * 60 * 1000);
  }

  /**
  * @summary Helper method for converting an Excel date value to a JavaScript timestamp
  */

  excelDateValueToTimestamp(excelDateValue: number) {
    return Math.round(this.daysToMilliseconds(excelDateValue - this.timeFormatDifference));
  }
  
  /**
  * @summary Constructs a JavaScript Date object from an Excel date value
  * @argument {number} excelDateValue Excel date value
  */

  constructor(excelDateValue: number) { 
    super();
    this.setTime(this.excelDateValueToTimestamp(excelDateValue));
  }
  
  /**
  * @summary Returns the date in Excel date format
  */

  toExcelDateValue() {
    return this.millisecondsToDays(this.getTime()) + this.timeFormatDifference;
  }

  /**
  * @summary Returns the date in the ISO 8601 format
  * @example 2024-12-01T12:30:00.000Z
  */
  
  toString(): string {
    return super.toISOString();
  }

  /**
  * @summary Returns the date in the US long format
  * @example December 1, 2024
  */
  
  toUSLongFormatString(): string {
    return super.toLocaleDateString('en-US', {
      month: 'long', day: 'numeric', year: 'numeric',
    });
  }

}
/**
* @summary ExcelURL class
* The ExcelURL class extends the standard JavaScript URL object.
*/

class ExcelURL extends URL {

  title: string;

  /**
  * @summary Constructs a JavaScript URL object from an Excel hyperlink
  * @argument {ExcelScript.RangeHyperlink} hyperlink Excel hyperlink
  */

  constructor(hyperlink: ExcelScript.RangeHyperlink) {
    super(hyperlink.address);
    this.title = hyperlink.textToDisplay;
  }

  /**
  * @summary Constructs an Excel hyperlink
  */

  toExcelHyperlink(): ExcelScript.RangeHyperlink {
    return {
      address: this.href,
      screenTip: this.title,
      textToDisplay: this.title
    }
  }

}
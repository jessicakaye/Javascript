**
 * Function converts column letter to column number
 * 
 * - @param columnLetter is required.
 * - @param startAt0or1 is optional.
 * 
 * - If @param startAt0or1 is 1,
 *   letter A is converted to number 1.
 * - If @param startAt0or1 is anything other than 1,
 *   letter A is converted to number 0.
 * 
 * - Google Sheets column limit is ZZZ
 *   (https://support.google.com/drive/answer/37603).
 */

function columnLetterToNumber(columnLetter, startAt0or1) {

  /**
   * Declare all variables
   */

  var alphabet;
  var alphabetLength;

  var columnNumber;

  var letterLength;
  var letterLeft;
  var letterMiddle;
  var letterRight;



  /**
   * Initialize main variables
   */

  alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  alphabetLength = alphabet.length; /*26*/
  letterLength = columnLetter.length;

  /**
   * Calculate column number
   */

  /* Get rightmost letter */
  letterRight = columnLetter.slice(letterLength - 1, letterLength);

  /* Convert rightmost letter to number */
  columnNumber = alphabet.indexOf(letterRight);


  if (letterLength > 1) {

    /* Get middle letter */
    letterMiddle = columnLetter.slice(letterLength - 2, letterLength - 1);

    /* Convert middle letter to number, and add it to existing column number */
    columnNumber = columnNumber + (alphabet.indexOf(letterMiddle) + 1) * alphabetLength;

    if (letterLength > 2) {
      /* Get leftmost letter */
      letterLeft = columnLetter.slice(letterLength - 3, letterLength - 2);
      /* Convert leftmost letter to number, and add it to existing column number */
      columnNumber = columnNumber + (alphabet.indexOf(letterLeft) + 1) * alphabetLength * alphabetLength;
    }
  }

  /**
   * Define if column numbers start at 0 or 1 
   */

  if (startAt0or1 == 1) {
    columnNumber++;
  }

  /**
   * Return column number
   */

  return columnNumber;
}


// hiding columns for P+ launch view
function hide_columns(columnletter) 
{ 
  var columnletter;
  //declaring required variables
  var d_out = new Array();
  // to access the required sheet names
  var name_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // use for loop to read name for all the available sheets
  for (var i=0 ; i<name_sheet.length ; i++) {
    if(name_sheet[i].getName().includes("(US")){d_out.push( [ name_sheet[i].getName() ] )};
  }
  // return statement for function definition 
//   return d_out 
Logger.log(d_out);

const ss = SpreadsheetApp.getActive();

  for( var j = 0; j < d_out.length; j++ ) {

    var current = ss.getSheetByName(d_out[j]);
    // //Hide sheets not in list
    Logger.log(d_out[j]);
    if (d_out[j].toString().includes("Paid Churn")) {
      var startcolumn = (columnLetterToNumber(columnletter, 1) - 1);
      Logger.log("True");
    }
    else {
      var startcolumn = columnLetterToNumber(columnletter, 1);
      Logger.log("False");

    }
    var endcolumn = startcolumn + 30;
    Logger.log(startcolumn);

    current.showColumns(2,endcolumn);
    current.hideColumns(1,1);
    current.hideColumns(startcolumn, 30);

  }
  
}

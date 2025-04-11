
// define function header
function show_source_sheets() 
{ 
  //declaring required variables
  var d_out = new Array();
  // to access the required sheet names
  var name_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // use for loop to read name for all the available sheets
  for (var i=0 ; i<name_sheet.length ; i++) {
    if(name_sheet[i].getName().includes("(1)")){d_out.push( [ name_sheet[i].getName() ] )}; // always keep this
    if(name_sheet[i].getName().includes("Validation")){d_out.push( [ name_sheet[i].getName() ] )};
    if(name_sheet[i].getName().includes("BQ")){d_out.push( [ name_sheet[i].getName() ] )};
  }
Logger.log(d_out);
// //Hide sheets not in list
  const ss = SpreadsheetApp.getActive();
  const list = d_out.flat();//assumed only one column
  ss.getSheets().filter(sh => ~list.indexOf(sh.getName())).forEach(sh => sh.showSheet()); //show every sheet
  ss.getSheets().filter(sh => !~list.indexOf(sh.getName())).forEach(sh => sh.hideSheet()); //hide every sheet

}

// define function header
function show_all_sheets() 
{ 
  //declaring required variables
  var d_out = new Array();
  // to access the required sheet names
  var name_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // use for loop to read name for all the available sheets
  for (var i=0 ; i<name_sheet.length ; i++) {
    d_out.push( [ name_sheet[i].getName() ] ) //chooses every single sheet
  }
  // return statement for function definition 
//   return d_out 
Logger.log(d_out);
// //Hide sheets not in list
  const ss = SpreadsheetApp.getActive();
  const list = d_out.flat();//assumed only one column
  ss.getSheets().filter(sh => ~list.indexOf(sh.getName())).forEach(sh => sh.showSheet()); //show every sheet

}


//deletes extra historical sheets
function delete_sheets() 
{ 
    // Display a dialog box with a message and "Yes" and "No" buttons.
  // The user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("WARNING!",
      "Are you sure you want to continue? All previous historical sheets will NOW be deleted. If unsure, click NO now.",
      ui.ButtonSet.YES_NO
  );

  // Process the user's response.
  if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes."');
  //declaring required variables
  var d_out = new Array();
  // to access the required sheet names
  var name_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // use for loop to read name for all the available sheets
  for (var i=0 ; i<name_sheet.length ; i++) {
    if(name_sheet[i].getName().includes("[TO_DELETE]")){d_out.push( [ name_sheet[i].getName() ] )};
Logger.log(d_out);
// //delete sheets in list
  const ss = SpreadsheetApp.getActive();
  const list = d_out.flat();//assumed only one column
  ss.getSheets().filter(sh => ~list.indexOf(sh.getName())).forEach(sh => ss.deleteSheet(sh)); //delete every sheet
    }
  }
  else {
      Logger.log(
          'The user clicked "No" or the close button in the dialog\'s title bar.'
      );
  }
}

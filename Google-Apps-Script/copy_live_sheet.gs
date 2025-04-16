
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function make_copy() 
    {    
        
  // Display a dialog box with a message and "Yes" and "No" buttons.
  // The user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Make Copy Now?",
      "Do you want to make a copy of the Live tab?",
      ui.ButtonSet.YES_NO
  );

  // Process the user's response.
  if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes."');
    
    var sheet = SpreadsheetApp.getActive()

  //refer to sheet with connected data
    var source = sheet.getSheetByName("LIVE")
   
  //create copy of connected data sheet
    source.copyTo(sheet);
    var spreadsheet = sheet.getSheetByName("Copy of LIVE");

  //rename
    var date = Utilities.formatDate(new Date(), sheet.getSpreadsheetTimeZone(), 'MM/dd/YY');
    spreadsheet.setName(date)
  
  // Copy formatting and values from Connected Starts sheet
    source.getRange("A1:CW500").copyTo(spreadsheet.getRange("A1:CW500"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,
false);


  }
  else {
      Logger.log(
          'The user clicked "No" or the close button in the dialog\'s title bar.'
      );
  }
}

  function copyConnectedTotalSubs() 
    {    
        
  // Display a dialog box with a message and "Yes" and "No" buttons.
  // The user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("WARNING! - Total Sub HHs",
      "Are you sure you want to continue? The previous historical sheet will be deleted. If unsure, click NO now.",
      ui.ButtonSet.YES_NO
  );

  // Process the user's response.
  if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes."');
  
    var sheet = SpreadsheetApp.getActive()

  //refer to sheets with connected data
    var source = sheet.getSheetByName("[Sub] Total Sub HHs Data (Connected)")
  //refer to historical sheet
    var destination = sheet.getSheetByName("[Sub] Total Sub HHs Data (Historical)");
   
  //create copy of connected data sheet
    source.copyTo(sheet);
    var spreadsheet = sheet.getSheetByName("Copy of [Sub] Total Sub HHs Data (Connected)");

  // create copy of historical just in case
    destination.copyTo(sheet);
    var date = Utilities.formatDate(new Date(), sheet.getSpreadsheetTimeZone(), 'MM-dd-yyyy');
    historicalcopy = sheet.getSheetByName("Copy of [Sub] Total Sub HHs Data (Historical)")
    historicalcopy.setName('[TO_DELETE] Total Sub HHs Data (Historical)' + " " + date)

  // Copy formatting and values from Connected Total Hours sheet
    destination.getRange("A1:CW500").copyTo(spreadsheet.getRange("A1:CW500"),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
false);
    source.getRange("A1:CW500").copyTo(spreadsheet.getRange("A1:CW500"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,
false);

    sheet.deleteSheet(destination);
    spreadsheet.setName("[Sub] Total Sub HHs Data (Historical)");



  //resets formulas on Total Hours Validation page with new historical data
  let textFinder = sheet
    .createTextFinder("=")
    .matchEntireCell(false)
    .matchCase(true)
    .matchFormulaText(true)
    .ignoreDiacritics(false)
    .replaceAllWith("=");
  
  // //hide data sheets
  // source.hideSheet();
  // spreadsheet.hideSheet();
  // historicalcopy.hideSheet();
  // sheet.getSheetByName("[Sub] Validation: Total Sub HHs").hideSheet();
  }
  else {
      Logger.log(
          'The user clicked "No" or the close button in the dialog\'s title bar.'
      );
  }
}

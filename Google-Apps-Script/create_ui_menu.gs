//Creates custom menu for BI purposes
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu("BI: Organization")
    .addItem("Choose Columns to Hide", "columnsPrompt")
    .addItem("Update Topline KPIs", "updatetoplineKPIs")
    .addItem("Order Sheets", "sortSheets")

    .addToUi();
}

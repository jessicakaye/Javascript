
//Creates custom prompt for knowing which tab to create images from
function tabPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
    "Print Tables",
    "From which tab to print?",
    ui.ButtonSet.OK_CANCEL,
  );

  // Process the user's response.
  var button = result.getSelectedButton();

  if (button == ui.Button.OK) {
    // User clicked "OK".
    // ui.alert("Printing tables from" + text + ".");
    var text = result.getResponseText();

    var result2 = ui.alert(
      "Print Tables",
      "Is this for a full month?",
      ui.ButtonSet.YES_NO,
  );


  if (result2 == ui.Button.YES) {
    // User clicked "OK".
    // ui.alert("Printing tables from" + text + ".");
    var text2 = 1;
  }
  else {
    var text2 = 0;
  }
  
  sheetToImage(text,text2);

  }

}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// async function sheetToImage(dt_today="04/01/25", mtd=1) {
async function sheetToImage(dt_today, mtd) {

// Display a dialog box with a message and "Yes" and "No" buttons.
  // The user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Print Tables?",
      "Do you want to print tables for " + dt_today + "?",
      ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    // Open the Google Sheet by ID
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName(dt_today); // Change to your sheet name

    if (mtd == 1) {
      var starts_range = sheet.getRange('b64:g76'); // Adjust the range as needed
      var subs_range = sheet.getRange('b78:g90'); // Adjust the range as needed
      var hrs_range = sheet.getRange('b92:g104'); // Adjust the range as needed
    }
    else {
      var starts_range = sheet.getRange('b19:g31'); // Adjust the range as needed
      var subs_range = sheet.getRange('b33:g45'); // Adjust the range as needed
      var hrs_range = sheet.getRange('b47:g59'); // Adjust the range as needed
    }

    // Get the range as an image
    // var imageBlob_starts = starts_range.build().getAs('image/png');
    // var imageBlob_subs = subs_range.build().getAs('image/png');
    // var imageBlob_hrs = hrs_range.build().getAs('image/png');

  // Specify the folder ID where you want to save the image
    var folder = DriveApp.getFolderById('1uL7coqClOquFzQIguYFF6IXMBowBJ69C');

    // Define the PDF export parameters
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?';
    var params = {
      exportFormat: 'pdf',
      format: 'pdf',
      size: 'A4',
      portrait: true,
      fitw: true,
      sheetnames: false,
      printtitle: false,
      pagenumbers: false,
      gridlines: false,
      fzr: false,
      range: starts_range.getA1Notation(),
      gid: sheet.getSheetId()
    };

    // Construct the URL with parameters
    var queryString = [];
    for (var param in params) {
      queryString.push(param + '=' + params[param]);
    }
    var exportUrl = url + queryString.join('&');

    // Fetch the PDF content
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    // Create a new file in Google Drive
    var pdfBlob_starts = response.getBlob().setName(sheet.getName() + '_starts.pdf');
    var file = folder.createFile(pdfBlob_starts);

      // Define the PDF export parameters

    var params = {
      exportFormat: 'pdf',
      format: 'pdf',
      size: 'A4',
      portrait: true,
      fitw: true,
      sheetnames: false,
      printtitle: false,
      pagenumbers: false,
      gridlines: false,
      fzr: false,
      range: subs_range.getA1Notation(),
      gid: sheet.getSheetId()
    };

    // Construct the URL with parameters
    var queryString = [];
    for (var param in params) {
      queryString.push(param + '=' + params[param]);
    }
    var exportUrl = url + queryString.join('&');

    // Fetch the PDF content
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    var pdfBlob_subs = response.getBlob().setName(sheet.getName() + '_subs.pdf');
    var file2 = folder.createFile(pdfBlob_subs);

    // Define the PDF export parameters
    var params = {
      exportFormat: 'pdf',
      format: 'pdf',
      size: 'A4',
      portrait: true,
      fitw: true,
      sheetnames: false,
      printtitle: false,
      pagenumbers: false,
      gridlines: false,
      fzr: false,
      range: hrs_range.getA1Notation(),
      gid: sheet.getSheetId()
    };

    // Construct the URL with parameters
    var queryString = [];
    for (var param in params) {
      queryString.push(param + '=' + params[param]);
    }
    var exportUrl = url + queryString.join('&');

    // Fetch the PDF content
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
    var pdfBlob_hrs = response.getBlob().setName(sheet.getName() + '_hrs.pdf');
    var file3 = folder.createFile(pdfBlob_hrs);


    // Create a new file in Google Drive
    // var file = DriveApp.createFile(imageBlob_starts.setName(sheet.getName() + '_starts.png'));
    // var file2 = DriveApp.createFile(imageBlob_subs.setName(sheet.getName() + '_subs.png'));
    // var file3 = DriveApp.createFile(imageBlob_hrs.setName(sheet.getName() + '_hrs.png'));

    // // Log the URL of the created image
    Logger.log('PDF URL: ' + file.getUrl());
    Logger.log('PDF URL: ' + file2.getUrl());
    Logger.log('PDF URL: ' + file3.getUrl());



    // Retrieve PDF data.
    // const fileId = file.getID() // Please set file ID of PDF file on Google Drive.
    const blob = file.getBlob();

    const pdfName = file.getName().replace('.pdf', ''); // Get the PDF name without the extension

    // Use a method for converting all pages in a PDF file to PNG images.
    const imageBlobs = await convertPDFToPNG_(blob,pdfName);

    // As a sample, create PNG images as PNG files.
    imageBlobs.forEach((b) => folder.createFile(b));


    // Retrieve PDF data.
    // const fileId = file.getID() // Please set file ID of PDF file on Google Drive.
    const blob2 = file2.getBlob();

    const pdfName2 = file2.getName().replace('.pdf', ''); // Get the PDF name without the extension

    // Use a method for converting all pages in a PDF file to PNG images.
    const imageBlobs2 = await convertPDFToPNG_(blob2,pdfName2);

    // As a sample, create PNG images as PNG files.
    imageBlobs2.forEach((b) => folder.createFile(b));



    // Retrieve PDF data.
    // const fileId = file.getID() // Please set file ID of PDF file on Google Drive.
    const blob3 = file3.getBlob();

    const pdfName3 = file3.getName().replace('.pdf', ''); // Get the PDF name without the extension

    // Use a method for converting all pages in a PDF file to PNG images.
    const imageBlobs3 = await convertPDFToPNG_(blob3,pdfName3);

    // As a sample, create PNG images as PNG files.
    imageBlobs3.forEach((b) => folder.createFile(b));


    file.setTrashed(true);
    file2.setTrashed(true);
    file3.setTrashed(true);


    ui.alert("Tables printed.");
  }

}
  
  
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * This is a method for converting all pages in a PDF file to PNG images.
 * PNG images are returned as BlobSource[].
 * IMPORTANT: This method uses Drive API. Please enable Drive API at Advanced Google services.
 *
 * @param {Blob} blob Blob of PDF file.
 * @return {BlobSource[]} PNG blobs.
 */
async function convertPDFToPNG_(blob,pdfName) {
  // Convert PDF to PNG images.
  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText()); // Load pdf-lib
  const setTimeout = function (f, t) {
    // Overwrite setTimeout with Google Apps Script.
    Utilities.sleep(t);
    return f();
  };
  const data = new Uint8Array(blob.getBytes());
  const pdfData = await PDFLib.PDFDocument.load(data);
  const pageLength = pdfData.getPageCount();
  console.log(`Total pages: ${pageLength}`);
  const obj = { imageBlobs: [], fileIds: [] };
  for (let i = 0; i < pageLength; i++) {
    console.log(`Processing page: ${i + 1}`);
    const pdfDoc = await PDFLib.PDFDocument.create();
    const [page] = await pdfDoc.copyPages(pdfData, [i]);

    // Crop the page (adjust the values as needed)
    const cropBox = { x: 35, y: 605, width: 520, height: 190 }; // Example crop box
    page.setCropBox(cropBox.x, cropBox.y, cropBox.width, cropBox.height);


    pdfDoc.addPage(page);
    const bytes = await pdfDoc.save();
    const blob = Utilities.newBlob(
      [...new Int8Array(bytes)],
      MimeType.PDF,
      `sample${i + 1}.pdf`
    );
    const id = DriveApp.createFile(blob).getId();
    Utilities.sleep(3000); // This is used for preparing the thumbnail of the created file.
    const link = Drive.Files.get(id, { fields: "thumbnailLink" }).thumbnailLink;
    if (!link) {
      throw new Error(
        "In this case, please increase the value of 3000 in Utilities.sleep(3000), and test it again."
      );
    }
    const imageBlob = UrlFetchApp.fetch(link.replace(/\=s\d*/, "=s1000"))
      .getBlob()
      .setName(`${pdfName}.png`);
    obj.imageBlobs.push(imageBlob);
    obj.fileIds.push(id);
  }
  obj.fileIds.forEach((id) => DriveApp.getFileById(id).setTrashed(true));
  return obj.imageBlobs;
}

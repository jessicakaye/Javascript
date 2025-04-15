
function copyTableToDoc() {

var ui = SpreadsheetApp.getUi();

// Prompt for the Doc ID
var docPrompt = ui.prompt(
  'Enter the Google Doc ID',
  'Paste everything after /d/ from the Doc URL',
  ui.ButtonSet.OK_CANCEL
);

if (docPrompt.getSelectedButton() !== ui.Button.OK) {
  ui.alert('Operation cancelled.');
  return;
}
var docId = docPrompt.getResponseText();

// Prompt for the sheet name
var sheetPrompt = ui.prompt(
  'Enter the tab (sheet) name',
  'Enter the name of the sheet you want to pull data from (e.g. "04/11/25")',
  ui.ButtonSet.OK_CANCEL
);

if (sheetPrompt.getSelectedButton() !== ui.Button.OK) {
  ui.alert('Operation cancelled.');
  return;
}
var sheetName = sheetPrompt.getResponseText();

  // Get the header prefix from the specified cell
  var headerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
  var headerPrefix = headerSheet.getRange("B10").getValue();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    // Open the existing Google Doc
  try {
    var doc = DocumentApp.openById(docId);
  } catch (e) {
    Logger.log('Error opening document: ' + e.message);
    return;
  }
  var body = doc.getBody();
  // Show a confirmation dialog 
  // var response = ui.alert('Confirm current quarter', 'The current quarter is: ' + headerPrefix + '. Do you want to proceed?', ui.ButtonSet.OK_CANCEL);
  
  // // Process the user's response
  // if (response == ui.Button.CANCEL) {
  //   ui.alert('Operation cancelled.');
  //   return;
  // }


  // Define the ranges and their headers
  var rangesAndHeaders = [
    { range: "A1:B4", header: "Active Sub HH Share By Platform" } //,
    //add more ranges here
  ];

// Track where tables are inserted
  const tableStartIndexes = [];

  body.appendParagraph(headerPrefix + " | Overall P+ Information")
    .setBold(true).setItalic(false).setForegroundColor("#0000FF").setUnderline(true);
  body.appendParagraph("");

    if (rangeInfo.header === "Active Sub HHs by Genre") {
      body.appendParagraph(""); body.appendParagraph(""); body.appendParagraph("");
      body.appendParagraph(headerPrefix + " | Detailed Genre & Content Type Information")
        .setBold(true).setForegroundColor("#0000FF").setUnderline(true);
      body.appendParagraph("");
    }

    const range = sheet.getRange(rangeInfo.range);
    const values = range.getDisplayValues();
    const backgrounds = range.getBackgrounds();
    const fonts = range.getFontFamilies();
    const fontSizes = range.getFontSizes();
    const fontColors = range.getFontColors();
    const fontWeights = range.getFontWeights();
    const fontStyles = range.getFontStyles();
    const horizontalAlignments = range.getHorizontalAlignments();

    // Header above the table
    body.appendParagraph(headerPrefix + " | " + rangeInfo.header)
      .setBold(true).setForegroundColor("#000000").setUnderline(false);


    // Capture the table start index before appending
    const startIndex = body.getChild(body.getNumChildren() - 1).getParent().getText().length;
    tableStartIndexes.push({ range: rangeInfo.range, startIndex: body.getText().length });

    // Append table
    const table = body.appendTable();
    for (let rowIdx = 0; rowIdx < values.length; rowIdx++) {
      const row = table.appendTableRow();
      for (let colIdx = 0; colIdx < values[rowIdx].length; colIdx++) {
        const cell = row.appendTableCell(values[rowIdx][colIdx]);
        cell.setBackgroundColor(backgrounds[rowIdx][colIdx]);
        const text = cell.editAsText();
        text.setFontFamily(fonts[rowIdx][colIdx]);
        text.setFontSize(fontSizes[rowIdx][colIdx]);
        text.setForegroundColor(fontColors[rowIdx][colIdx]);
        text.setBold(fontWeights[rowIdx][colIdx] === 'bold');
        text.setItalic(fontStyles[rowIdx][colIdx] === 'italic');
        text.setUnderline(false); // <-- THIS LINE REMOVES UNDERLINE
        try {
          cell.getChild(0).asParagraph().setAlignment(
            DocumentApp.HorizontalAlignment[horizontalAlignments[rowIdx][colIdx].toUpperCase()]
          );
        } catch (e) {
          Logger.log('Alignment issue: ' + e.message);
        }
      }
    }
  }

  body.appendParagraph(""); body.appendParagraph(""); body.appendParagraph("");
  body.appendParagraph(headerPrefix + " | Title-Level Information")
    .setBold(true).setForegroundColor("#0000FF").setUnderline(true);
  body.appendParagraph("");

  doc.saveAndClose(); // Ensure body updates are saved so Docs API sees them

  // Merge cells using Docs API — example: first table, merge first row, columns 0-1
 const mergeRequests = [];
const docData = Docs.Documents.get(docId);
const content = docData.body.content;

let contentIndex = 0;
let tableIdx = 0;

for (let i = 0; i < rangesAndHeaders.length; i++) {
  const rangeInfo = rangesAndHeaders[i];
  const range = sheet.getRange(rangeInfo.range);
  const mergedRanges = range.getMergedRanges();
  
  // Find the i-th table in the document content
  while (contentIndex < content.length && !content[contentIndex].table) contentIndex++;
  if (contentIndex >= content.length) break; // no more tables found

  const tableStartIndex = content[contentIndex].startIndex;

  // Convert each merged range from Sheets into Docs API mergeTableCells request
  mergedRanges.forEach(mergedRange => {
    const startRow = mergedRange.getRow() - range.getRow();
    const startCol = mergedRange.getColumn() - range.getColumn();
    const rowSpan = mergedRange.getNumRows();
    const colSpan = mergedRange.getNumColumns();

    mergeRequests.push({
      mergeTableCells: {
        tableRange: {
          tableCellLocation: {
            tableStartLocation: {
              index: tableStartIndex
            },
            rowIndex: startRow,
            columnIndex: startCol
          },
          rowSpan: rowSpan,
          columnSpan: colSpan
        }
      }
    });
  });

  contentIndex++; // move to next table
}


  if (mergeRequests.length > 0) {
    Docs.Documents.batchUpdate(
      { requests: mergeRequests },
      docId
    );
  }

  ui.alert('Tables copied and cells merged!');
}

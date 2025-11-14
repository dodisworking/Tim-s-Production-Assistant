function doGet() {
  // Return the HTML dialog for the web app
  return HtmlService.createHtmlOutputFromFile('Dialog')
    .setTitle('tester')
    .setWidth(400)
    .setHeight(200);
}

function doPost(e) {
  // Handle form submission
  var value = e.parameter.inputValue;
  writeToCell(value);
  return HtmlService.createHtmlOutput('<html><body><h2>Value written to A1!</h2><script>window.close();</script></body></html>');
}

function writeToCell(value) {
  // Replace with your Spreadsheet ID, or use getActiveSpreadsheet() if running from the sheet
  var spreadsheetId = 'YOUR_SPREADSHEET_ID_HERE'; // Replace with your actual spreadsheet ID
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('TimeSlots');
  
  // Alternative: if the script is deployed from the sheet itself, use:
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TimeSlots');
  
  if (sheet) {
    sheet.getRange('A1').setValue(value);
  } else {
    throw new Error('Sheet "TimeSlots" not found!');
  }
}


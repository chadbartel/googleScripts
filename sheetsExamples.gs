// Reading information from a sheet
function logProductInfo() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    Logger.log('Product name: ' + data[i][0]);
    Logger.log('Product number: ' + data[i][1]);
  }
}

// Writing information to a sheet
function addProduct() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['Cotton Sweatshirt XL', 'css004']);
}

// Formatting data on a sheet
function formatMySpreadsheet() {
  // Set the font style of the cells in the range of B2:C2 to be italic.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cell = sheet.getRange('B2:C2');
  cell.setFontStyle('italic');
}

// Validating data on a sheet
function validateMySpreadsheet() {
  // Set a rule for the cell B4 to be a number between 1 and 100.
  var cell = SpreadsheetApp.getActive().getRange('B4');
  var rule = SpreadsheetApp.newDataValidation()
     .requireNumberBetween(1, 100)
     .setAllowInvalid(false)
     .setHelpText('Number must be between 1 and 100.')
     .build();
  cell.setDataValidation(rule);
}

// Generate a chart from sheet data
function newChart() {
  // Generate a chart representing the data in the range of A1:B15.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.BAR)
     .addRange(sheet.getRange('A1:B15'))
     .setPosition(5, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
}

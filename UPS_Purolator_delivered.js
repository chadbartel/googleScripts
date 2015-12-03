function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Get Shipment Confirmation", functionName: "getShipmentConfirm"});
  ss.addMenu("Run", menuEntries);
}

function getShipmentConfirm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mk = ss.getSheetByName("Maint Kit Service Requests");
  var range = mk.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  
  for (var i=1; i<=numRows; i++) {
    for (var j=1; j<=numCols; j++) {
      var currentValue = range.getCell(i, j).getValue().toString();
      if (currentValue.length === 18) {
        var row = range.getCell(i, j).getRow();
        var col = range.getCell(i, j).getColumn();
        var UPSConfirm = getUPSConfirm(row, col);
        range.getCell(i, j).offset(0, 1).setValue(UPSConfirm);
      // This 'if' statement is not executing
      // DEBUG
      } else if (currentValue.length === 12) {
        var row = range.getCell(i, j).getRow();
        var col = range.getCell(i, j).getColumn();
        var PuroConfirm = getPuroConfirm(row, col);
        range.getCell(i, j).offset(0, 1).setValue(PuroConfirm);
      // The above statement is being skipped and executing below
      } else {
        console.error("Row " + i + " doesn't have a valid value.");
      }
    }
  }
}

function getUPSConfirm(r, c) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var si = ss.getSheetByName("Shipping Info");
  var mk = ss.getSheetByName("Maint Kit Service Requests");
  var formulaCell = si.getRange(1, 1);
  var inputCell = mk.getRange(r, c);
  var resultCell = si.getRange(765, 1);
  var formulaStr = "=IMPORTDATA(\"http://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=\"&'Maint Kit Service Requests'!";   
  
  si.clearContents();
  formulaCell.setValue(formulaStr.concat(inputCell.getA1Notation(), ")"));

  while (true) {
    if (resultCell.getValue() != "") break
  }
  
  return isDelivered(resultCell.getValue());
  
}

function getPuroConfirm(r, c) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var si = ss.getSheetByName("Shipping Info");
  var mk = ss.getSheetByName("Maint Kit Service Requests");
  var formulaCell = si.getRange(1, 1);
  var inputCell = mk.getRange(r, c);
  var resultCell = si.getRange(1176, 1);
  var formulaStr = "=IMPORTDATA(\"http://www.purolator.com/purolator/ship-track/tracking-details.page?pin=\"&'Maint Kit Service Requests'!";   
  
  si.clearContents();
  formulaCell.setValue(formulaStr.concat(inputCell.getA1Notation(), ")"));

  while (true) {
    if (resultCell.getValue() != "") break
  }
  
  return isDelivered(resultCell.getValue());
  
}

function isDelivered(cellVal) {
  // determine whether the mk has been delivered
  if (cellVal === "Delivered") {
    return "Yes";
  } else if (cellVal === "status: 'Delivered'") {
    return "Yes";
  } else {
    return "";
  }
}

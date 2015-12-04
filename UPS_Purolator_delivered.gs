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
      } else if (currentValue.length === 12) {
        var row = range.getCell(i, j).getRow();
        var col = range.getCell(i, j).getColumn();
        var PuroConfirm = getPuroConfirm(row, col);
        range.getCell(i, j).offset(0, 1).setValue(PuroConfirm);
      } else {
        Logger.log("Row " + i + " doesn't have a valid value.");
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
  var formulaStr = "=IMPORTDATA(\"http://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=\"&'Maint Kit Service Requests'!";   
  var pageData = si.getRange(2, 1, si.getMaxRows()-1, 1);
  var numRows = pageData.getNumRows();
  
  si.clearContents();
  formulaCell.setValue(formulaStr.concat(inputCell.getA1Notation(), ")"));
  
  var checkPage = checkPageLoad();
  while (checkPage != true) {
    Logger.log("Loading page...");
    checkPage = checkPageLoad();
  }
  
  for (var i = 2; i <= numRows; i++) {
    switch (pageData.getCell(i, 1).getValue()) {
      case "Arrival Scan":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "At Local Post Office":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Delivered":
        Logger.log(pageData.getCell(i, 1).getValue());
        return "Yes";
      case "Departure Scan":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Destination Scan":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Exception Action Required":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Export Scan":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Given to Post Office for Delivery":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Import Scan":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Order Processed: In Transit to UPS":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Order Processed: Ready for UPS":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "Origin Scan":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "On Vehicle for Delivery":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      default:
        continue;
    }
  }
  Logger.log("No tracking found.")
  return "";
}

function getPuroConfirm(r, c) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var si = ss.getSheetByName("Shipping Info");
  var mk = ss.getSheetByName("Maint Kit Service Requests");
  var formulaCell = si.getRange(1, 1);
  var inputCell = mk.getRange(r, c);
  var formulaStr = "=IMPORTDATA(\"http://www.purolator.com/purolator/ship-track/tracking-details.page?pin=\"&'Maint Kit Service Requests'!";   
  var pageData = si.getRange(2, 1, si.getMaxRows()-1, 1);
  var numRows = pageData.getNumRows();
  
  si.clearContents();
  formulaCell.setFormula(formulaStr.concat(inputCell.getA1Notation(), ")"));
  
  var checkPage = checkPageLoad();
  while (checkPage != true) {
    Logger.log("Loading page...");
    checkPage = checkPageLoad();
  }
  
  for (var i = 2; i <= numRows; i++) {
    switch (pageData.getCell(i, 1).getValue()) {
      case "status: 'Received by Purolator'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Corrective Action - Currently In Transit'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Delivered'":
        Logger.log(pageData.getCell(i, 1).getValue());
        return "Yes";
      case "status: 'Shipment In Transit'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'On Purolator vehicle for delivery'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Attempted Delivery- Customer Closed'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Address Correction Required'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Attempted Delivery - Package Refused'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Attempted Delivery- Receiver Unavailable'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Scheduled Delivery Appointment Required'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Customer Requested PM Delivery'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Pending Customs Clearance'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Delayed due to Weather'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Mechanical Delay - Currently In Transit'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Delivery Rescheduled'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Delayed - Incomplete Shipment'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'In Transit in U.S.'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'No Scanning Detail Available'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      case "status: 'Invalid Tracking Number Entered - Please Re-enter.'":
        Logger.log(pageData.getCell(i, 1).getValue());
        break;
      default:
        continue;
    }
  }
  Logger.log("No tracking found.");
  return "";
}

function checkPageLoad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var si = ss.getSheetByName("Shipping Info");
  var pageData = si.getRange(2, 1, si.getMaxRows()-1, 1);
  var numRows = pageData.getNumRows();
  
  while (true) {
    for (var i = 2; i <= numRows; i++) {
      var currentValue = pageData.getCell(i, 1).getValue();
      if (currentValue === "") {
        Logger.log("Searching for data...");
      } else {
        break;
      }
      break;
    }
    break;
  }
  
  return true;
}

function onEdit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var editedRow = ss.getActiveCell().getRowIndex();
  var editorEmail = Session.getActiveUser().getEmail();
  if (checkRange(editedRow)) {
    if (ss.getRange(editedRow, 11).getValue()) {
      //do nothing
    } else {
      ss.getRange(editedRow, 11).setValue(editorEmail);
    }
  } else {
    ss.getRange(editedRow, 11).clear();
  }
}

function checkRange(editRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  for (var i = 1; i <= 10; i++) {
    cVal = ss.getRange(editRow, i).getValue();
    if (cVal == "") {
      // do nothing and check next value
    } else {
      return true;
    }
  }
  return false;
}

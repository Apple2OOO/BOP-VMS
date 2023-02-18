function findPreviousDayIn() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Sign-In");
  var auditSheet = SpreadsheetApp.getActive().getSheetByName("Audit");

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[5]);
    if (timestamp.toDateString() == yesterday.toDateString() && row[7] == "In") {
      auditSheet.appendRow([row[3]]);
    }
  }
}
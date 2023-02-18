function dailyAudit() {
  
var shtMR = SpreadsheetApp.getActive().getSheetByName("Macro References");
var lastRowAudit = shtMR.getRange('G3:G').getValues().filter(String).length + 2;
var cellsAuditNames = shtMR.getRange('G3:G'+lastRowAudit);
var cellsAuditPoints = shtMR.getRange('J3:J'+lastRowAudit);

var shtEE = SpreadsheetApp.getActive().getSheetByName("Event Entry");
var lastRowEE = shtEE.getRange('C:C').getValues().filter(String).length + 4;
var formattedDate = Utilities.formatDate(new Date(), "EST", "M/d/yyyy");
var formattedTimestamp = Utilities.formatDate(new Date(), "EST", "EEE MMM dd yyyy HH:mm:ss");

shtEE.getRange('A'+lastRowEE).setValue('VMS Audit: Day of ' + formattedDate + '. Conducted on '+ formattedTimestamp);
cellsAuditNames.copyTo(shtEE.getRange('B'+lastRowEE), {contentsOnly: true});
cellsAuditPoints.copyTo(shtEE.getRange('C'+lastRowEE), {contentsOnly: true});
}
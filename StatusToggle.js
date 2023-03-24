function directorToggle() {

  // Open VMS, Navigate to "Macro References" tab and extract VMS Dashboard URL from cell B6
  var dashboardUrl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Macro References').getRange(6,2).getValue();

  // Open VMS Dashboard using dashboardUrl. Store as wbDashboard variable
  var wbDashboard = SpreadsheetApp.openByUrl(dashboardUrl);
  
  // Get Director indicator from VMS Dashboard on "Dashboard" tab cell G13
  var indicatorCell = wbDashboard.getSheetByName('Dashboard').getRange(13,7);

  // If indicator == in, change to out. If indicator == out, change to in. Else, change to out.
  if(indicatorCell.getValue() == "Out") {
    indicatorCell.setValue('In');
  } else if (indicatorCell.getValue() == "In"){
    indicatorCell.setValue('Out');
  } else {
    indicatorCell.setValue('Out');
  }
}

function asstDirectorToggle() {

  // Open VMS, Navigate to "Macro References" tab and extract VMS Dashboard URL from cell B6
  var dashboardUrl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Macro References').getRange(6,2).getValue();

  // Open VMS Dashboard using dashboardUrl. Store as wbDashboard variable
  var wbDashboard = SpreadsheetApp.openByUrl(dashboardUrl);
  
  // Get Director indicator from VMS Dashboard on "Dashboard" tab cell G14
  var indicatorCell = wbDashboard.getSheetByName('Dashboard').getRange(14,7);

  // If indicator == in, change to out. If indicator == out, change to in. Else, change to out.
  if(indicatorCell.getValue() == "Out") {
    indicatorCell.setValue('In');
  } else if (indicatorCell.getValue() == "In"){
    indicatorCell.setValue('Out');
  } else {
    indicatorCell.setValue('Out');
  }
}


function onOpen() {
 SpreadsheetApp.getUi()
 .createMenu("Forms")
  .addItem("Request Audit", "showSidebar")
  .addItem("Missed Visitor", "MissedVisitor")
  .addToUi();
}

function showSidebar() {
 SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Sidebar.html").setTitle("Audit Request Form"));
}

function MissedVisitor() {
 SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("MissedVisitor.html").setTitle("Missed Visitor"));
}
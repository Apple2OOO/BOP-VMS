//@OnlyCurrentDoc

function onOpen() {
 SpreadsheetApp.getUi().createMenu("Audit").addItem("Request Audit", "showSidebar").addToUi();
}

function showSidebar() {
 SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Sidebar.html").setTitle("Audit Request Form"));
}
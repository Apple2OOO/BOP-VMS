function SortScoreboard() {
  var shtScoreboard = SpreadsheetApp.getActive().getSheetByName("Scoreboard");
  shtScoreboard.getFilter().sort(3, false);
};
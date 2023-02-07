function onFormSubmit(e) 
{
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    var formulaDisplayName ='=IF(ISBLANK(D'+lastRow+'),"",TRIM(C'+lastRow+')&", "&TRIM(B'+lastRow+'))';
    var formulaLookup = '=IF(ISBLANK(D'+lastRow+'),"",IF(ISNUMBER(MATCH(VALUE(D'+lastRow+'),\'Members List\'!A:A,0)),"","ADD"))';
    sheet.getRange(lastRow, 7).setFormula(formulaDisplayName);
    sheet.getRange(lastRow, 8).setFormula(formulaLookup);
    var newUserID = sheet.getRange(lastRow, 4).getValue();
    var newUserDisplayName = sheet.getRange(lastRow, 7).getValue();
    var newUserFirstName = sheet.getRange(lastRow, 2).getValue();
    var newUserLastName = sheet.getRange(lastRow, 3).getValue();
    var newUserComboKey = newUserFirstName + newUserLastName;
    var newUserEmail = sheet.getRange(lastRow, 5).getValue();

  if(sheet.getRange(lastRow,8).getValue() == "ADD") 
  {
    var shtMembers = e.source.getSheetByName('Members List');
    var lastRowMembers = shtMembers.getRange('B1:B').getValues().filter(String).length + 1;
    shtMembers.getRange('A' + lastRowMembers).setValue(newUserID); 
    shtMembers.getRange('B' + lastRowMembers).setValue(newUserDisplayName);
    shtMembers.getRange('C' + lastRowMembers).setValue(newUserFirstName);
    shtMembers.getRange('D' + lastRowMembers).setValue(newUserLastName);
    shtMembers.getRange('E' + lastRowMembers).setValue(newUserComboKey);
    shtMembers.getRange('F' + lastRowMembers).setValue(newUserID);
    shtMembers.getRange('H' + lastRowMembers).setValue(newUserEmail);
  }
}
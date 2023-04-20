function onFormSubmit(e) 
{
    var sheet = e.range.getSheet();

  if(sheet.getName() == "New Users"){
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
} else if (sheet.getName() == "Missed Visitors"){
      var lastRow = sheet.getLastRow();

    //    var formulaLookup = '=IF(ISBLANK(D'+lastRow+'),"",IF(ISNUMBER(MATCH(VALUE(D'+lastRow+'),\'Members List\'!A:A,0)),"","ADD"))';
      //  sheet.getRange(lastRow, 8).setFormula(formulaLookup);
        var visitTimeStamp = sheet.getRange(lastRow, 1).getValue();
        var visitorFirstName = sheet.getRange(lastRow, 2).getValue();
        var visitorLastName = sheet.getRange(lastRow, 3).getValue();
        var visitorEmail = sheet.getRange(lastRow, 4).getValue();
        var visitorMeeting = sheet.getRange(lastRow, 5).getValue();
        var visitPurpose = sheet.getRange(lastRow, 6).getValue();
        var visitReturn = sheet.getRange(lastRow, 7).getValue();
        var currentSecretary = sheet.getRange(lastRow, 8).getValue();
        var notes = sheet.getRange(lastRow, 9).getValue();
    // Browser.msgBox('Starting');
      sendEmail(visitTimeStamp,visitorFirstName, visitorLastName,visitorEmail,visitorMeeting,visitPurpose,visitReturn,currentSecretary,notes);
    // Browser.msgBox('Finished');
}

function sendEmail(visitTimeStamp,visitorFirstName, visitorLastName,visitorEmail,visitorMeeting,visitPurpose,visitReturn,currentSecretary,notes) {
if (visitorMeeting == 'Darren'){ var recipient = 'henrydl@purdue.edu'}
else if (visitorMeeting == 'Darien') {var recipient = 'thomp347@purdue.edu'}
else if (visitorMeeting == 'Darren, Darien') {var recipient = ['henrydl@purdue.edu, thomp347@purdue.edu']}
  var visitTimeStamp = Utilities.formatDate(visitTimeStamp, "EST", "EEE MMM dd yyyy HH:mm:ss");
  //var recipient = 'appletoj@purdue.edu';
  var subject = "Missed Visitor in BOP Office";
  var body = 'Missed Visitor on ' + visitTimeStamp + '\n\n' + 
              'Visitor Name: '+ visitorFirstName + ' ' + visitorLastName + '\n' +
              'Visitor Email: ' + visitorEmail +  '\n\n' +
              'Who were they here to meet?: ' + visitorMeeting + '\n' +
              'Purpose of Visit: '+ visitPurpose + '\n\n'+
              'When were they instructed to return? '+ visitReturn + '\n' + 
              'Current Student Secretary: ' + currentSecretary + '\n\n' +
              'Notes: ' + notes;



  GmailApp.sendEmail(recipient, subject, body);
}

}
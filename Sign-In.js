function onEdit(e) {

  var sheet = e.source.getActiveSheet(); // Get Active Sheet
  var row = e.range.getRow(); // Get Entry Row
  var col = e.range.getColumn(); // Get Entry Column
  if(sheet.getName() == "Sign-In"){

  /*if(col != 1) {
    var editPassword = Browser.inputBox('Enter Password to Edit this Cell');
    if(editPassword == 'admin'){
      return;
    } else{
      Browser.msgBox('Invalid Password');

    }
  }*/

  var eventType = sheet.getRange('L6').getValue() // Get Current Event
  var sc = e.source.getSheetByName('Scoreboard') // Store Scoreboard Sheet


// ---- RESET
// Check if column 1, non-header row, Sign-In Sheet, blank ID
  if(col == 1 && row > 1 && e.source.getActiveSheet().getName() == "Sign-In" && sheet.getRange(row,1).getValue() == "") {
    sheet.getRange(row+1,1).activate(); // Activate next cell down
  sheet.getRange('A' + row  + ":" + 'K' + row).setValue(""); // Clear values
    sheet.getRange('A' + row  + ":" + 'K' + row).setBackground(null); // Clear blackouts
     sheet.getRange(row,1).activate() // Return ActiveCell to deleted row
  //sc.getFilter().sort(3, false); // Sort Scoreboard
  return;
  }
// ---- Enter Data
  // Check edit is in column 1, non-header row, sign-in sheet, and not blank
  if(col == 1 && row > 1 && e.source.getActiveSheet().getName() == "Sign-In" && sheet.getRange(row,1).getValue() != "") { 
      addFormulas(sheet,row,eventType);
  }
// ---- Check Unregistered
  if(col == 1 && row > 1 && e.source.getActiveSheet().getName() == "Sign-In" && sheet.getRange(row,3).getValue() == "Not Registered") 
  {
  Browser.msgBox('ID Not Found. Please Have Visitor Register Using QR Code. Delete this entry and ask visitor to sign in again after completing the form.')//Register Now?')//, Browser.Buttons.YES_NO); 
   sheet.getRange(A + row).setValue("H");
  /* if(newUserRegistration == "yes")
   {
     var shtMembers = e.source.getSheetByName('Members List');
  var newUserFirstName = Browser.inputBox('Enter Your First Name');
  var newUserLastName = Browser.inputBox('Enter Your Last Name');
  var newUserEmail = Browser.inputBox('Enter Your Purdue Email');
  var newUserID = sheet.getRange('A' + row).getValue();

  var lastRowMembers = shtMembers.getRange('B1:B').getValues().filter(String).length + 1;
  shtMembers.getRange('A' + lastRowMembers).setValue(newUserID); 
  shtMembers.getRange('B' + lastRowMembers).setValue(newUserLastName + ', ' + newUserFirstName);
      // --- Add Visitor Formula
      e.source.getSheetByName('Sign-In').getRange('C' + row + '').setFormula('=ifna(IF(ISBLANK($A' + row + ')=TRUE,TRIM(""),VLOOKUP($B' + row + ',\'Members List\'!A:B,2,0)),"Not Registered")');
   }*/
  }


  }  

// SECRETARY
  // ---- Enter Data
  // Check edit is in column 1, non-header row, sign-in sheet, and not blank
   if(col == 1 && row > 1 && e.source.getActiveSheet().getName() == "Secretary") {
    //Browser.msgBox('D');
    logSecretary(sheet,row,col);
   } 



}
function addFormulas(sheet, row, eventType) {
    sheet.getRange(row,4).setValue(eventType); // Set Event Type
        sheet.getRange(row,6).setValue(new Date()); // Set Timestamp
        sheet.getRange(row,1).setBackgroundColor('#000000'); // Black out ID

      //Add Formulas to Row
      // --- Add PUID Formula
        sheet.getRange('B' + row + '').setFormula('=IF(OR(LEN($A' + row + ')=8,LEN($A' + row + ')=10),$A' + row + ',IF(ISBLANK($A' + row + ')=TRUE,TRIM(""),VALUE(MID($A' + row + ',19,8))))');

      // --- Add Visitor Formula
        sheet.getRange('C' + row + '').setFormula('=ifna(IF(ISBLANK($A' + row + ')=TRUE,TRIM(""),VLOOKUP($B' + row + ',\'Members List\'!A:B,2,0)),"Not Registered")');
      // --- In/Out Status Column  
               sheet.getRange('G' + row).setFormula('=if(countif($C$2:$C' + row + ',C' + row + ')-1=0,"In",if(I' + row + '="No Out","In",if(isblank(A' + row + '),"",if(isblank(H' + row + '),"",if(if(isblank(A' + row + '),"",H' + row + ')="Out","In","Out")))))');
      // --- Last In/Out Column
        sheet.getRange('H' + row).setFormula('=if((I' + row + ')="No Out","",(LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(G$2:G' + (row-1) + ',A$2:A' + (row-1) + ',TRUE))))\n');
      // --- Last Timestamp Column
        sheet.getRange('I' + row + '').setFormula('=if(isblank(A' + row + '),"",if(day((LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(F$2:F' + (row-1) + ',A$2:A' + (row-1) + ',TRUE))))<>day(F' + row + '),"No Out",""))\n');
      // --- Duration Column
        sheet.getRange('J' + row + '').setFormula('=if(G' + row + '<>"Out","", if(isblank(A' + row + '),"",abs(((LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(F$2:F' + (row-1) + ',A$2:A' + (row-1) + ',TRUE)))-$F' + row + '))))');

  sheet.getRange('E' + row).setFormula('=if(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))<=1),VLOOKUP(D'+row+',\'Point Types\'!$A:$B,2,0),IF(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))>1),0,If(G'+row+'="Out",IF(AND(((24*60)*(J'+row+'))>=15,((24*60)*(J'+row+'))<=360),ROUND((J'+row+'*1440)*(\'Macro References\'!$B$2)),0),"")))');

      // Break Visitor Formula
          sheet.getRange(row,3).setValue(sheet.getRange(row,3).getValue());

      // Return ActiveCell to Next Blank Row
        sheet.getRange(row+1,1).activate(); 
}

function setSignInCell() {
  var shtSignIn = SpreadsheetApp.getActive().getSheetByName('Sign-In');
  var lastrowSignIn = shtSignIn.getRange('A1:A').getValues().filter(String).length + 1;
  Browser.msgBox(lastrowSignIn);
}

function logSecretary(sheet,row,col) {
      sheet.getRange(row,4).setValue('Secretary'); // Set Event Type
      sheet.getRange(row,6).setValue(new Date()); // Set Timestamp
      sheet.getRange(row,1).setBackgroundColor('#000000'); // Black out ID


// ---- RESET
// Check if column 1, non-header row, Sign-In Sheet, blank ID
  if(sheet.getRange(row,1).getValue() == "") {
    sheet.getRange(row+1,1).activate(); // Activate next cell down
  sheet.getRange('A' + row  + ":" + 'K' + row).setValue(""); // Clear values
    sheet.getRange('A' + row  + ":" + 'K' + row).setBackground(null); // Clear blackouts
     sheet.getRange(row,1).activate() // Return ActiveCell to deleted row
 // sc.getFilter().sort(3, false); // Sort Scoreboard
  }
  else if(sheet.getRange(row,1).getValue() != "") {

      //Add Formulas to Row
      // --- Add PUID Formula
        sheet.getRange('B' + row + '').setFormula('=IF(OR(LEN($A' + row + ')=8,LEN($A' + row + ')=10),$A' + row + ',IF(ISBLANK($A' + row + ')=TRUE,TRIM(""),VALUE(MID($A' + row + ',19,8))))');

      // --- Add Visitor Formula
        sheet.getRange('C' + row + '').setFormula('=ifna(IF(ISBLANK($A' + row + ')=TRUE,TRIM(""),VLOOKUP($B' + row + ',\'Members List\'!A:B,2,0)),"Not Registered")');
      // --- In/Out Status Column  
       // sheet.getRange('G' + row).setFormula('=if(I' + row + '="No Out","In",if(isblank(A' + row + '),"",if(isblank(H' + row + '),"",if(if(isblank(A' + row + '),"",H' + row + ')="Out","In","Out"))))\n');
               sheet.getRange('G' + row).setFormula('=if(countif($C$2:$C' + row + ',C' + row + ')-1=0,"In",if(I' + row + '="No Out","In",if(isblank(A' + row + '),"",if(isblank(H' + row + '),"",if(if(isblank(A' + row + '),"",H' + row + ')="Out","In","Out")))))');

      // --- Last In/Out Column
        sheet.getRange('H' + row).setFormula('=if((I' + row + ')="No Out","",(LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(G$2:G' + (row-1) + ',A$2:A' + (row-1) + ',TRUE))))\n');
      // --- Last Timestamp Column
        sheet.getRange('I' + row + '').setFormula('=if(isblank(A' + row + '),"",if(day((LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(F$2:F' + (row-1) + ',A$2:A' + (row-1) + ',TRUE))))<>day(F' + row + '),"No Out",""))\n');
      // --- Duration Column
        sheet.getRange('J' + row + '').setFormula('=if(G' + row + '<>"Out","", if(isblank(A' + row + '),"",abs(((LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(F$2:F' + (row-1) + ',A$2:A' + (row-1) + ',TRUE)))-$F' + row + '))))');
      // --- Points Column
  //          sheet.getRange('E' + row + '').setFormula('=IF(C' + row + '="",TRIM(""),IF(G' + row + '="Out",ROUNDUP((J' + row + '*1440)*(\'Macro References\'!$B$2)),if(vlookup(C' + row + ',\'Daily Entries\'!A:C,3,0)=0,VLOOKUP(D' + row + ',\'Point Types\'!$A:$B,2,0),0)))')

 if(sheet.getName() == "Sign-In"){ 
 sheet.getRange('E' + row).setFormula('=if(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))<=1),VLOOKUP(D'+row+',\'Point Types\'!$A:$B,2,0),IF(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))>1),0,If(G'+row+'="Out",IF(AND(((24*60)*(J'+row+'))>=15,((24*60)*(J'+row+'))<=360),ROUND((J'+row+'*1440)*(\'Macro References\'!$B$2)),0),"")))');
 } else if(sheet.getName() == "Secretary"){
sheet.getRange('E' + row).setFormula('=if(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))<=1),VLOOKUP(D'+row+',\'Point Types\'!$A:$B,2,0),IF(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))>1),0,If(G'+row+'="Out",IF(AND(((24*60)*(J'+row+'))>=15,((24*60)*(J'+row+'))<=360),ROUND((J'+row+'*1440)*(\'Macro References\'!$B$3)),0),"")))');
 }



      // Break Visitor Formula
          sheet.getRange(row,3).setValue(sheet.getRange(row,3).getValue());

      // Return ActiveCell to Next Blank Row
        sheet.getRange(row+1,1).activate(); 
  }
}
function setCell() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCell = sheet.getRange("A" + (lastRow + 1));
  sheet.setActiveRange(lastCell);
}

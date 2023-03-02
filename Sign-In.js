function onEdit(e) {

  var sheet = e.source.getActiveSheet(); // Get Active Sheet
  var row = e.range.getRow(); // Get Entry Row
  var col = e.range.getColumn(); // Get Entry Column
  if(sheet.getName() == "Sign-In"){

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
  }


  }  

// SECRETARY
  // ---- Enter Data
  // Check edit is in column 1, non-header row, sign-in sheet, and not blank
   if(col == 1 && row > 1 && e.source.getActiveSheet().getName() == "Secretary") {
    //Browser.msgBox('D');
    logSecretary(sheet,row,col);
   } 

 SortScoreboard();

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
               sheet.getRange('G' + row).setFormula('=if(countif($C$2:$C' + row + ',C' + row + ')-1=0,"In",if(I' + row + '="No Out","In",if(isblank(A' + row + '),"",if(isblank(H' + row + '),"",if(if(isblank(A' + row + '),"",H' + row + ')="Out","In","Out")))))');

// --- Last In/Out Column
        sheet.getRange('H' + row).setFormula('=if((I' + row + ')="No Out","",(LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(G$2:G' + (row-1) + ',A$2:A' + (row-1) + ',TRUE))))\n');
// --- Last Timestamp Column
        sheet.getRange('I' + row + '').setFormula('=if(isblank(A' + row + '),"",if(day((LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(F$2:F' + (row-1) + ',A$2:A' + (row-1) + ',TRUE))))<>day(F' + row + '),"No Out",""))\n');
// --- Duration Column
        sheet.getRange('J' + row + '').setFormula('=if(G' + row + '<>"Out","", if(isblank(A' + row + '),"",abs(((LOOKUP(A' + row + ',sort(A$2:A' + (row-1) + '),SORT(F$2:F' + (row-1) + ',A$2:A' + (row-1) + ',TRUE)))-$F' + row + '))))');
// --- Points Column
 if(sheet.getName() == "Sign-In"){ 
 sheet.getRange('E' + row).setFormula('=if(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))<=1),VLOOKUP(D'+row+',\'Point Types\'!$A:$B,2,0),IF(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))>1),0,If(G'+row+'="Out",IF(AND(((24*60)*(J'+row+'))>=15,((24*60)*(J'+row+'))<=360),ROUND((J'+row+'*1440)*(\'Macro References\'!$B$2)),0),"")))');
 } else if(sheet.getName() == "Secretary"){
sheet.getRange('E' + row).setFormula('=if(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))<=1),VLOOKUP(D'+row+',\'Point Types\'!$A:$B,2,0),IF(AND(G'+row+'="In",COUNTIFS(C$2:C'+row+',C'+row+',F$2:F'+row+',">="&int(F'+row+'))>1),0,If(G'+row+'="Out",IF(AND(((24*60)*(J'+row+'))>=15,((24*60)*(J'+row+'))<=360),ROUND((J'+row+'*1440)*(\'Macro References\'!$B$3)),0),"")))');

//-- Log user out on Sign-In tab
/*
if( User is logged in ) */
  var shtSignIn = SpreadsheetApp.getActive().getSheetByName('Sign-In'); // Get Sign-In Sheet
  var secID = sheet.getRange('A' + row).getValue(); // Get ID from Sec Sheet
  var secSISOlastRow = shtSignIn.getLastRow(); // Get Last Row of SI Sheet
  var secSISOlastCell = shtSignIn.getRange('A' + (secSISOlastRow + 1)); // Get First Available Cell in SI sheet
  
  
  //secSISOlastCell.setValue(secID); // Add Sec ID to SI sheet
  lookupID(shtSignIn, secID);


// Will need a lookup here to check if user is signed in

function lookupID(shtSignIn, secID) {
  
  var sheet = shtSignIn;
  var data = shtSignIn.getDataRange().getValues();
  var secID = secID; // replace with the actual ID value you want to lookup
  
  for (var i = data.length-1; i >= 0; i--) {
    var lookupRow = data[i];
    var ID = lookupRow[0];
    var status = lookupRow[6];
    var passRow = data.length;
    
    if (ID == secID && status == "In") {
      //var ui = SpreadsheetApp.getUi();
      //Browser.msgBox("Please sign out");
      secretarySignOutFormulas(sheet,passRow);
     // Browser.msgBox("Through");
      //var result = ui.alert("Please sign out", ui.ButtonSet.OK);
      return;
    }
  }
}



// Now do we loop back through the formulas again or go to its own separate tree?


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

function secretarySignOutFormulas(sheet,row)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sign-In');
 // Browser.msgBox(row);
 // Browser.msgBox(sheet.getName());
  //Add Formulas to Row
Browser.msgBox('Please make sure to sign out on the \'Sign-In\' Sheet before starting your shift');

}
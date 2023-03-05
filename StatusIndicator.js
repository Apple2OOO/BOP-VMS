function statusIndicatorDirector() {
  var wb = SpreadsheetApp.getActive();
  var sht = wb.getActiveSheet();

  if(sht.getName() == "Dashboard") {
    var statusDirector = sht.getRange('G13').getValue();
    //var statusAsstDirector = sht.getRange('G14').getValue();

    if (statusDirector == "In") {
        sht.getRange('G13').setValue("Out");
    }
    else if(statusDirector == "Out"){
        sht.getRange('G13').setValue("In");
    }

  }

}

function statusIndicatorAsstDirector() {
  var wb = SpreadsheetApp.getActive();
  var sht = wb.getActiveSheet();

  if(sht.getName() == "Dashboard") {
    var statusAsstDirector = sht.getRange('G14').getValue();

    if (statusAsstDirector == "In") {
        sht.getRange('G14').setValue("Out");
    }
    else if(statusAsstDirector == "Out"){
        sht.getRange('G14').setValue("In");
    }

  }

}


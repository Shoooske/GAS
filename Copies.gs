//function onOpen(){
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var menus = [{name: 'Copy to OverviewTotal', functionName: 'copyToOverviewTotal'},
//               {name: 'Copy to OverviewConfirmed', functionName: 'copyToOverviewConfirmed'},
//               {name: 'Copy to OverviewDone', functionName: 'copyToOverviewDone'}
//              ];
//  ss.addMenu('Copy', menus);
//}

function copyToOverviewTotal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var sheetTo = ss.getSheetByName("OverviewTotal");
  var sheetToLength = sheetTo.getLastRow();
  var numColumns = 9;

  if (sheetToLength > 5) {
    sheetTo.deleteRows(6, sheetToLength - 5);
  }

  for (var i = 7; i < sheets.length; i++) {
    var sheetFrom = sheets[i];
    var dtFrom = sheetFrom.getDataRange().getValues();

    for (var row = 5; row <= dtFrom.length; row++) {
      if (dtFrom[row - 1][0] != "Invalidation" && dtFrom[row - 1][0] != "Done") {
        var toRow = setValue(dtFrom,sheetTo,row,numColumns);

        if (dtFrom[row - 1][0] == "Coming") {
          sheetTo.getRange(toRow, 1, 1, numColumns).setBackground("#ffff00");
        } else if (dtFrom[row - 1][0] == "Active") {
          sheetTo.getRange(toRow, 1, 1, numColumns).setBackground("#7cfc00");
        }
      }
    }
  }
}

function copyToOverviewAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var sheetTo = ss.getSheetByName("OverviewAll");
  var sheetToLength = sheetTo.getLastRow();
  var numColumns = 9;

  if (sheetToLength > 5) {
    sheetTo.deleteRows(6, sheetToLength - 5);
  }

  for (var i = 7; i < sheets.length; i++) {
    var sheetFrom = sheets[i];
    var dtFrom = sheetFrom.getDataRange().getValues();

    for (var row = 5; row <= dtFrom.length; row++) {
      if (dtFrom[row - 1][0] != "Invalidation") {
        setValue(dtFrom,sheetTo,row,numColumns);
      }
    }
  }
}


function copyToOverviewConfirmed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var sheetTo = ss.getSheetByName("OverviewConfirmed");
  var sheetToLength = sheetTo.getLastRow();
  var numColumns = 16;

  if (sheetToLength > 5) {
    sheetTo.deleteRows(6, sheetToLength - 5);
  }

  for (var i = 7; i < sheets.length; i++) {
    var sheetFrom = sheets[i];
    var dtFrom = sheetFrom.getDataRange().getValues();

    for (var row = 5; row <= dtFrom.length; row++) {
      if (dtFrom[row-1][0] == "Active" || dtFrom[row-1][0] == "Coming") {

        var toRow = setValue(dtFrom,sheetTo,row,numColumns);

        if (dtFrom[row - 1][0] == "Coming") {
          sheetTo.getRange(toRow, 1, 1, 16).setBackground("#ffff00");
        } else if (dtFrom[row - 1][0] == "Active") {
          sheetTo.getRange(toRow, 1, 1, 16).setBackground("#7cfc00");
        }

      }
    }
  }
}


function copyToOverviewDone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var sheetTo = ss.getSheetByName("OverviewDone");
  var sheetToLength = sheetTo.getLastRow();
  var numColumns = 8;

  if (sheetToLength > 5) {
    sheetTo.deleteRows(6, sheetToLength - 5);
  }

  for (var i = 7; i < sheets.length; i++) {
    var sheetFrom = sheets[i];
    var dtFrom = sheetFrom.getDataRange().getValues();

    for (var row = 5; row <= dtFrom.length; row++) {
      if (dtFrom[row - 1][0] == "Done") {
        setValue(dtFrom,sheetTo,row, numColumns);
      }
    }
  }
}


function setValue(dtFrom,sheetTo,row,numColumns) {
  //共通項目
  var copyFirstName = dtFrom[row - 1][1];
  var copyFamilyName = dtFrom[row - 1][2];
  var copyCompanyName = dtFrom[row - 1][4];
  var copySex = dtFrom[row - 1][8];
  var copyAge = dtFrom[row - 1][9];
  var copyDateOfStart = dtFrom[row - 1][22];
  var copyDateOfFinal = dtFrom[row - 1][23];
  var copyCourse = dtFrom[row - 1][21];

  if(numColumns == 16){
    var copyFlightArrDate = dtFrom[row - 1][14];
    var copyFlightArrTime = dtFrom[row - 1][15];
    var CopyFlightArrNo = dtFrom[row - 1][16];
    var copyFlightDepDate = dtFrom[row - 1][17];
    var copyFlightDepTime = dtFrom[row - 1][18];
    var copyFlightDepNo = dtFrom[row - 1][19];
    var copyArrStatus = dtFrom[row - 1][24];
  }

  //OverviewTotal , OverviewAll
  if(numColumns >= 9){
    var copyStudentStatus = dtFrom[row - 1][0];
  }

  var SampleFormat = sheetTo.getRange(5, 1, 1, numColumns);

  var dtTo = sheetTo.getDataRange().getValues();

  for (var toRow = 6; toRow <= dtTo.length + 1; toRow++) {

    if (sheetTo.getRange(toRow, 2).getValue() != "") {
      if (copyDateOfStart >= sheetTo.getRange(toRow, 7).getValue()) {
        continue;
      }
      sheetTo.insertRowBefore(toRow);
      SampleFormat.copyFormatToRange(sheetTo, 1, numColumns, toRow, toRow);
    }

    sheetTo.getRange(toRow, 2).setValue(copyFirstName);
    sheetTo.getRange(toRow, 3).setValue(copyFamilyName);
    sheetTo.getRange(toRow, 4).setValue(copyCompanyName);
    sheetTo.getRange(toRow, 5).setValue(copySex);
    sheetTo.getRange(toRow, 6).setValue(copyAge);
    sheetTo.getRange(toRow, 7).setValue(copyDateOfStart);
    sheetTo.getRange(toRow, 8).setValue(copyDateOfFinal);
    sheetTo.getRange(toRow, 9).setValue(copyCourse);

    if(numColumns == 16){
      sheetTo.getRange(toRow, 10).setValue(copyFlightArrDate);
      sheetTo.getRange(toRow, 11).setValue(copyFlightArrTime);
      sheetTo.getRange(toRow, 12).setValue(CopyFlightArrNo);
      sheetTo.getRange(toRow, 13).setValue(copyFlightDepDate);
      sheetTo.getRange(toRow, 14).setValue(copyFlightDepTime);
      sheetTo.getRange(toRow, 15).setValue(copyFlightDepNo);
      sheetTo.getRange(toRow, 16).setValue(copyArrStatus);
      sheetTo.getRange(toRow, 1).setValue(copyStudentStatus);
    }

    if(numColumns == 9){
      sheetTo.getRange(toRow, 1).setValue(copyStudentStatus);
    }
    return toRow;
  }
}


//function copyToOverviewDone() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//
//  var sheetFrom = ss.getActiveSheet();
//  var dtFrom = sheetFrom.getDataRange().getValues();
//
//  var sheetTo = ss.getSheetByName("OverviewDone");
//
//  for (var row = 5; row <= dtFrom.length; row++) {
//    if (dtFrom[row-1][33] != "Done") {
//
//      var copyFirstName = dtFrom[row-1][0];
//      var copyFamilyName = dtFrom[row-1][1];
//      var copyCompanyName = dtFrom[row-1][3];
//      var copySex = dtFrom[row-1][7];
//      var copyAge = dtFrom[row-1][8];
//      var copyDateOfStart = dtFrom[row-1][21];
//      var copyDateOfFinal = dtFrom[row-1][22];
//      var copyCourse = dtFrom[row-1][20];
//
//      var dtTo = sheetTo.getDataRange().getValues();
//
//      for (var toRow = 5; toRow <= dtTo.length + 1; toRow++) {
//        if (sheetTo.getRange(toRow, 1).getValue() == "") {
//
//          sheetTo.getRange(toRow, 1).setValue(copyFirstName);
//          sheetTo.getRange(toRow, 2).setValue(copyFamilyName);
//          sheetTo.getRange(toRow, 3).setValue(copyCompanyName);
//          sheetTo.getRange(toRow, 4).setValue(copySex);
//          sheetTo.getRange(toRow, 5).setValue(copyAge);
//          sheetTo.getRange(toRow, 6).setValue(copyDateOfStart);
//          sheetTo.getRange(toRow, 7).setValue(copyDateOfFinal);
//          sheetTo.getRange(toRow, 8).setValue(copyCourse);
//
//          break;
//        } else if (sheetTo.getRange(toRow, 1).getValue() != "" && copyDateOfStart < sheetTo.getRange(toRow, 6).getValue()) {
//
//          sheetTo.insertRowBefore(toRow);
//
//          sheetTo.getRange(toRow, 1).setValue(copyFirstName);
//          sheetTo.getRange(toRow, 2).setValue(copyFamilyName);
//          sheetTo.getRange(toRow, 3).setValue(copyCompanyName);
//          sheetTo.getRange(toRow, 4).setValue(copySex);
//          sheetTo.getRange(toRow, 5).setValue(copyAge);
//          sheetTo.getRange(toRow, 6).setValue(copyDateOfStart);
//          sheetTo.getRange(toRow, 7).setValue(copyDateOfFinal);
//          sheetTo.getRange(toRow, 8).setValue(copyCourse);
//
//          break;
//        }
//
//      }
//
//    }
//  }
//}
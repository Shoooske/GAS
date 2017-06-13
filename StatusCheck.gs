//Add menu bar
function onOpen() {

  //menu array
  var myMenu=[
    {name: "Show connection status", functionName: "createAllSheet"},
    {name: "Map status", functionName: "mapDesks"}
  ];

  SpreadsheetApp.getActiveSpreadsheet().addMenu("Classroom", myMenu);
}


////Create new classroom sheet
//function createSheet() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getActiveSheet();
//  var sheetName = sheet.getName();
//  
//  var dt = sheet.getDataRange().getValues();
//
//  var newSs = SpreadsheetApp.create("sheetName_" + dt[3][8] + "_ConnectionStatus");
//  var newSheet = newSs.getActiveSheet();
//  
//  sheet.getRange(6, 11).setValue(newSs.getId());
//  newSheet.getRange(1, 1).setValue("Name");
//  newSheet.getRange(1, 2).setValue("ID");
//  newSheet.getRange(1, 3).setValue("Status");
//  
//  var row = 7;
//  while (dt[row][8] != "" && dt[row][8] != "undefined") {
//    if (dt[row][10] != "") {
//      newSheet.getRange(row,1).setValue(dt[row][10]);
//      newSheet.getRange(row,2).setValue(dt[row][9].slice(18));
//      if (dt[row][11] == 1) {
//        newSheet.getRange(row,3).setValue("Connected").setBackground('#b0c4de');
//      } else if (dt[row][11] == 0) {
//        newSheet.getRange(row,3).setValue("Offline").setBackground('#ffffe0');
//      } else {
//        newSheet.getRange(row,3).setValue("Error").setBackground('#ff0000');
//      }
//    }
//    row += 1;
//  }
//  
//  var newLastRow = newSheet.getLastRow();
//  
//  for(var i = newLastRow; i > 0 ; --i){
//    var cellA = newSheet.getRange(i,1).getValue();
//
//    if(cellA == "") {
//      newSheet.deleteRow(i);
//    }
//  }
//  
//  Browser.msgBox("You can view connection status of " + "dt[3][8]" + " ->" + newSs.getUrl());
//
//}


//Create new classroom sheet
function createAllSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  
  var dt = sheet.getDataRange().getValues();

  
  var row = 7;
  var column = 8;
  
  while (dt[3][column] != "" && dt[3][column] != undefined) {
    var newSs = SpreadsheetApp.create(sheetName + "_" + dt[3][column] + "_ConnectionStatus");
    var newSheet = newSs.getActiveSheet();
    
    sheet.getRange(6, column + 3).setValue(newSs.getId());
    newSheet.getRange(1, 1).setValue("Name");
    newSheet.getRange(1, 2).setValue("ID");
    newSheet.getRange(1, 3).setValue("Status");

    while (dt[row][column] != "" && dt[row][column] != undefined) {
      if (dt[row][column + 2] != "") {
        newSheet.getRange(row,1).setValue(dt[row][column + 2]);
        newSheet.getRange(row,2).setValue(dt[row][column + 1].slice(18));
        if (dt[row][column + 3] == 1) {
          newSheet.getRange(row,3).setValue("Connected").setBackground('#b0c4de');
        } else if (dt[row][column + 3] == 0) {
          newSheet.getRange(row,3).setValue("Offline").setBackground('#ffffe0');
        } else {
          newSheet.getRange(row,3).setValue("Error").setBackground('#ff0000');
        }
      }
      row += 1;
    }
    
    var newLastRow = newSheet.getLastRow();

    for(var i = newLastRow; i > 0 ; --i){
      var cellA = newSheet.getRange(i,1).getValue();
      if(cellA == "") {
        newSheet.deleteRow(i);
      }
    }
    
    Browser.msgBox("You can view connection status of " + dt[3][column] + " ->" + newSs.getUrl());
    column += 6;
    row = 7;
  }

}


//Desk mapping
function mapDesks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dt = sheet.getDataRange().getValues();
  
  
  var column = 8;
  
  while (dt[3][column] != "" && dt[3][column] != undefined) {
    var ID = sheet.getRange(6, column + 3).getValue();
    
    var newSs = SpreadsheetApp.openById(ID);
    var newSheet = newSs.getActiveSheet();
    
    var noOfStudents = dt[3][12]
    var deskColumn = dt[5][12]
    
    var i = 2;
    var j = 5;
    var k = 2;
    while(k <= 1 + noOfStudents) {
      var targetToCopy = newSheet.getRange(i, j);
      newSheet.getRange(k, 3).copyTo(targetToCopy);
      if(j < 5 + 2 * (deskColumn -1)) {
        j = j + 2;
      } else {
        i = i + 2;
        j = 5;
      }

      k += 1;

    }
    
    column += 6;
    i = 2;
    j = 5;
    k = 2;

  }  
}

//Update classroom sheet
function updateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dt = sheet.getDataRange().getValues();

  var column = 8;

  while (dt[3][column] != "" && dt[3][column] != undefined) {
    var ID = sheet.getRange(6, column + 3).getValue();
    
    var newSs = SpreadsheetApp.openById(ID);
    var newSheet = newSs.getActiveSheet();
    
    var row = 7;
    var rowNew = 2;
    while (dt[row][column] != "" && dt[row][column] != undefined) {
      if (dt[row][column + 2] != "") {
        newSheet.getRange(rowNew, 3)
        if (dt[row][column + 3] == 1) {
          newSheet.getRange(rowNew, 3).setValue("Connected").setBackground('#b0c4de');
        } else if (dt[row][column + 3] == 0) {
          newSheet.getRange(rowNew, 3).setValue("Offline").setBackground('#ffffe0');
        } else {
          newSheet.getRange(rowNew, 3).setValue("Error").setBackground('#ff0000');
        }
        rowNew += 1;
      }
      row += 1;
    }
    
    var map = newSheet.getRange(2, 5).getValue();
    if (map != "") {
      var noOfStudents = dt[3][12]
      var deskColumn = dt[5][12]
    
      var i = 2;
      var j = 5;
      var k = 2;
      while(k <= 1 + noOfStudents) {
        //var status = newSheet.getRange(k, 3).getValue();
        var targetToCopy = newSheet.getRange(i, j);
        newSheet.getRange(k, 3).copyTo(targetToCopy);
        if(j < 5 + 2 * (deskColumn -1)) {
          j = j + 2;
        } else {
          i = i + 2;
          j = 5;
        }
        
        k += 1;
        
      }
      
      column += 6;
      i = 2;
      j = 5;
      k = 2;
      
    }  
    
    column += 6;
    row = 7;
  }
 
}


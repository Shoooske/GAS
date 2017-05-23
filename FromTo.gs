function sheet_name() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("DeleteManager");

  var cnt = spreadsheet.getSheets().length;
  var cell = sheet.getRange("B2:B3");

  var name_array = [];
  
  for (var i = 1; i <= cnt; i++) {
    var arr_sh = spreadsheet.getSheets();
    var name = arr_sh[i-1].getName();
    
    name_array.push(name);
  }

  var rule = SpreadsheetApp
　  　　.newDataValidation()
　　  　.requireValueInList(name_array, true)
　　　  .build();

  　cell.setDataValidation(rule);
}

function delete_sheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("DeleteManager");
  var arr_sh = spreadsheet.getSheets();

  var FromSheet_name = sheet.getRange("B2").getValue();
  var FromIndex = spreadsheet.getSheetByName(FromSheet_name).getIndex();
  
  var ToSheet_name = sheet.getRange("B3").getValue();
  var ToIndex = spreadsheet.getSheetByName(ToSheet_name).getIndex(); 
  
  var cnt = ToIndex - FromIndex + 1;

  var selections = Browser.msgBox("選択したシートを削除しますか？", Browser.Buttons.YES_NO);
  if (selections == 'yes') {
    for(var i = FromIndex; i <= ToIndex; i++) {
      spreadsheet.deleteSheet(arr_sh[i-1]);
    }
    Browser.msgBox(cnt + "枚のシートが削除されました。");
  }
}


//
//function test() {
//  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = spreadsheet.getSheetByName("DeleteManager");
//    var arr_sh = spreadsheet.getSheets();
//
//  var FromSheet_name = sheet.getRange("E4").getValue();
//  var FromSheet = spreadsheet.getSheetByName(FromSheet_name);
//  var FromIndex = FromSheet.getIndex();
//  
//  Logger.log(FromIndex);
//}

//シート「DeleteManager」に各シートごとの最終更新日を記録
function LastUpdated() {
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var id = sheet.getIndex();
  var EditedUTC = new Date();
  
  var del_sheet = Spreadsheet.getSheetByName("DeleteManager");

  del_sheet.getRange(id+1, 2).setValue(EditedUTC);
}

//30日間更新のないシートの削除
function delete_old_sheets(){
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var del_sheet = Spreadsheet.getSheetByName("DeleteManager");
  
  var cnt = Spreadsheet.getSheets().length;
  var now = new Date();

  for (var i = cnt; i >= 1; i--) {
    var arr_sh = Spreadsheet.getSheets();
    var last_update = del_sheet.getRange(i+1, 2).getValue();
    var name = arr_sh[i-1].getName();
    
    if(now.getTime() > last_update.getTime() + 86400000*30) {
      var selections = Browser.msgBox("最終更新から30日経過しています。シート「" + name + "」を削除してもよろしいですか？", Browser.Buttons.YES_NO);
      if (selections == 'yes') {
        Spreadsheet.deleteSheet(arr_sh[i-1]);
        Browser.msgBox("このシートは削除されました。");
      }
    }
  }
}
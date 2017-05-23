//セルA1に最終更新日記入
function insertLastUpdated() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var EditedUTC = new Date();

  sheet.getRange('A1').setValue("Last Update");
  sheet.getRange('B1').setValue(EditedUTC);
}

//日付の入力スペースを空けるために行１を追加
function InsertRowBefore() {
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = Spreadsheet.getActiveSheet();
  var A1 = sheet.getRange('A1').getValue();
  var B1 = sheet.getRange('B1').getValue();
  
  var EditedUTC = new Date();
  
  if(A1 != "" || B1 != "") {
    sheet.insertRowBefore(1);
    insertLastUpdated();
  } else {
    insertLastUpdated();
  }
}

//古いシートの削除
function delete_obsolete_sheets(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cnt = sheet.getSheets().length;
  var now = new Date();

  for (var i = cnt; i >= 1; i--) {
    var arr_sh = sheet.getSheets();
    var date = arr_sh[i-1].getRange("B1").getValue();
    var name = arr_sh[i-1].getName();
    
    if(now.getTime() > date.getTime() + 86400000*30) {
      var selections = Browser.msgBox("最終更新から30日経過しています。シート「" + name + "」を削除しますか？", Browser.Buttons.OK_CANCEL);
      if (selections == 'ok') {
        sheet.deleteSheet(arr_sh[i-1]);
        Browser.msgBox("このシートは削除されました。");
      }
    }
  }
}

  

////var file = DriveApp.getFileById(docsId);
////var lastUpdated = file.getLastUpdated();

////シート数カウント
//function sheet_counts(){
//  var sheet = SpreadsheetApp.getActiveSpreadsheet();
//  var cnt = sheet.getSheets().length;
//
//  Browser.msgBox(cnt);
//}


////セルA1に最終更新日記入(JST)
//function insertLastUpdated() {
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var EdUTC = new Date();
//  var EdJST = Utilities.formatDate(EdUTC, 'Asia/Tokyo', 'yyyy年M月d日'); //JST変換
//
//  sheet.getRange('A1').setValue("最終更新日");
//  sheet.getRange('B1').setValue(EdJST);
//}

////古いシートの削除(アクティブシートでのテスト)
//function delete_obsolete_sheets(){
//  var sheet = SpreadsheetApp.getActiveSpreadsheet();
//  var cnt = sheet.getSheets().length;
//  var now = new Date();
//  var nowJST = now.toString();
//
//    var arr_sh = sheet.getActiveSheet();
//    var date = arr_sh.getRange("B1").getValue();
//    var name = arr_sh.getName();
//    
//    if(now.getTime() > date.getTime() + 86400000*30) {
//      var selections = Browser.msgBox("最終更新から30日経過しています。シート「" + name + "」を削除しますか？", Browser.Buttons.OK_CANCEL);
//      if (selections == 'ok') {
//        sheet.deleteSheet(arr_sh[i]);
//        Browser.msgBox("このシートは削除されました。");
//      }
//    }
//
//}

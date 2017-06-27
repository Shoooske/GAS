/**
 * 留学チーム用　統計ツール
 *
 * シートの順番は固定なので、ずらさないように。
 *
 *
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menus = [{name: 'All Companies', functionName: 'updateAllCompanies'},
    {name: 'Only THIS company', functionName: 'updateOneCompany'},
  ];
  ss.addMenu('Updates', menus);
}

function updateOneCompany() {
  updateOneCompanyStudentStatus();
  copyToOverviewTotal();
  copyToOverviewConfirmed();
  copyToOverviewDone();
  copyToOverviewAll();  
}

function updateAllCompanies() {
  updateAllCompaniesStudentStatus();
  copyToOverviewTotal();
  copyToOverviewConfirmed();
  copyToOverviewDone();
  copyToOverviewAll();  
}
               
               

//On Company Sheet
/**
 * dt[row-1][0] 生徒のステータス
 * dt[row-1][22] 開始日
 * dt[row-1][23] 終了日
 */
function updateOneCompanyStudentStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  updateStudentStatus(sheet);
}


//On All Company Sheets
function updateAllCompaniesStudentStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  for (var i = 7; i < sheets.length; i++) {
    updateStudentStatus(sheets[i]);
  }
}

function updateStudentStatus(sheet){
  var dataTable = sheet.getDataRange().getValues();
  var now = new Date();

  for (var row = 5; row <= dataTable.length; row++) {
    if (dataTable[row-1][0] == "Booking" && dataTable[row-1][22] <= now) {
      sheet.getRange(row, 1).setValue("Invalidation");
    } else if (dataTable[row-1][0] == "Coming" && dataTable[row-1][22] <= now) {
      sheet.getRange(row, 1).setValue("Active");
    } else if (dataTable[row-1][0] == "Active" && dataTable[row-1][23] <= now) {
      sheet.getRange(row, 1).setValue("Done");
    }
  }
}


//function updateStudentStatusOnOverviewTotal() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getSheetByName("OverviewTotal");
//  var dt = sheet.getDataRange().getValues();
//
//  var now = new Date();
//
//  for (var row = 6; row <= dt.length; row++) {
//    if (dt[row-1][8] == "Booking" && dt[row-1][5] <= now) {
//      sheet.getRange(row, 9).setValue("Invalidation");
//    } else if (dt[row-1][8] == "Coming" && dt[row-1][5] <= now) {
//      sheet.getRange(row, 9).setValue("Active");
//      sheet.getRange(row, 1, 1, 9).setBackground("#7cfc00");
//    } else if (dt[row-1][8] == "Active" && dt[row-1][6] <= now) {
//      sheet.getRange(row, 9).setValue("Done");
//      sheet.getRange(row, 1, 1, 9).setBackground("#ffffff");
//    }
//  }
//}


//function updateStudentStatusOnOverviewConfirmed() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getSheetByName("OverviewConfirmed");
//  var dt = sheet.getDataRange().getValues();
//
//  var now = new Date();
//
//  for (var row = 6; row <= dt.length; row++) {
//    if (dt[row-1][15] == "Booking" && dt[row-1][5] <= now) {
//      sheet.getRange(row, 15).setValue("Invalidation");
//    } else if (dt[row-1][15] == "Coming" && dt[row-1][5] <= now) {
//      sheet.getRange(row, 15).setValue("Active");
//      sheet.getRange(row, 1, 1, 16).setBackground("#7cfc00");
//    } else if (dt[row-1][15] == "Active" && dt[row-1][6] <= now) {
//      sheet.getRange(row, 15).setValue("Done");
//      sheet.getRange(row, 1, 1, 16).setBackground("#ffffff");
//    }
//  }
//}
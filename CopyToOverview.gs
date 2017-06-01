//Add menu bar
function onOpen() {

  //menu array
  var myMenu = [
    {name: "Post to Overview", functionName: "copyValues"}
  ];

  SpreadsheetApp.getActiveSpreadsheet().addMenu("Coordination", myMenu);
}


/**
 * Transcription from SLCC to Overview
 */
function copyValues() {
  //Variable of SLCC
  var sheetFrom = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  /** @var Array  All Data from Original Sheet */
  var dtFrom = sheetFrom.getDataRange().getValues();
  var dtFromLength = dtFrom.length;

  //Variable of Overview
  var ssTo = SpreadsheetApp.openById('1Z0bXcsiSDwybq_6ueZUQ0dheoRCOMrer6sPkRexAPnM');
  var sheetTo = ssTo.getSheetByName('Overview');

  /** @var Array  All Data from Overview */
  var dtTo = sheetTo.getDataRange().getValues();
  var dtToLength = dtTo.length;

  //Execution of Posting
  //Executing Row by Row from SLCC
  for (var i = 10; i <= dtFromLength; i++) {
    if (dtFrom[i - 1][10] == "2. Feasibility Confirmed"
      && dtFrom[i - 1][11] != "1. Yes") { //Execute under These Conditions
      
      copyDate = dtFrom[i - 1][5].getTime();

      //Screening Overview Row by Row
      for (var row = 299; row < dtToLength; row++) { //Overviewが現在298行目からのため。日付整形後に、299->2としたい（Asteriaの日付がイレギュラー）。
        if (dtTo[row][1] != ""
          && copyDate < dtTo[row][1].getTime()) {
          
          sheetTo.insertRowAfter(row);
          var SampleFormat = sheetTo.getRange(11, 1, 1, 27);
          SampleFormat.copyFormatToRange(sheetTo, 1, 27, row + 1, row + 1);

          //Updating "Input Date"
          sheetTo.getRange(row + 1, 1).setValue(new Date());

          //Transcription of "Lesson Date"
          var copyLessonDate = dtFrom[i - 1][5];
          sheetTo.getRange(row + 1, 2).setValue(copyLessonDate);
          
          //Updating "Lesson Day"
          var arrDay = new Array('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat');
          sheetTo.getRange(row + 1, 3).setValue(arrDay[copyLessonDate.getDay()]);

          //Transcription of "School Client"
          var copyClientName = dtFrom[1][4];
          var copyGrade = dtFrom[i - 1][2];
          var copyClass = dtFrom[i - 1][3];
          sheetTo.getRange(row + 1, 4).setValue(copyClientName + " " + copyGrade + "-" + copyClass); //Post to "Project Name" in This Style
          
          //Updating as "Regular"
          sheetTo.getRange(row + 1, 5).setValue("Regular");

          //Transcription of "No. Student"
          var copyNoOfStudents = dtFrom[i - 1][4];
          sheetTo.getRange(row + 1, 6).setValue(copyNoOfStudents);

          //Transcription of "Lesson Time(PHT)"
          var copyLessonTimePHT = dtFrom[i - 1][7];
          sheetTo.getRange(row + 1, 8).setValue(copyLessonTimePHT);

          //Updating "Status"
          sheetTo.getRange(row + 1, 11).setValue("Feasibility Confirmed");

          //Transcription of "Ph Ops Team Checker"
          var copyPhPic = dtFrom[3][4];
          sheetTo.getRange(row + 1, 12).setValue(copyPhPic);

          //Transcription of "Platform"
          var copyPlatform = dtFrom[4][4];
          sheetTo.getRange(row + 1, 13).setValue(copyPlatform);

          //Transcription of "Material"
          var copyMaterial = dtFrom[i - 1][8];
          sheetTo.getRange(row + 1, 16).setValue(copyMaterial);

          //Transcription of "Customize Pattern"
          var copyCustomize = dtFrom[i - 1][9];
          sheetTo.getRange(row + 1, 18).setValue(copyCustomize);

          //Transcription of "JP Sales Team PP"
          var copyJpnPic = dtFrom[2][4];
          sheetTo.getRange(row + 1, 23).setValue(copyJpnPic);

          //Updating "In Overview Tab?"
          sheetFrom.getRange(i, 12).setValue("1. Yes");

          break;
        } else if (dtTo[row][0] == "") {
          
          sheetTo.insertRowAfter(row);
          var SampleFormat = sheetTo.getRange(11, 1, 1, 27);
          SampleFormat.copyFormatToRange(sheetTo, 1, 27, row + 1, row + 1);

          //Updating "Input Date"
          sheetTo.getRange(row + 1, 1).setValue(new Date());

          //Transcription of "Lesson Date"
          var copyLessonDate = dtFrom[i - 1][5];
          sheetTo.getRange(row + 1, 2).setValue(copyLessonDate);
          
          //Updating "Lesson Day"
          var arrDay = new Array('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat');
          sheetTo.getRange(row + 1, 3).setValue(arrDay[copyLessonDate.getDay()]);

          //Transcription of "School Client"
          var copyClientName = dtFrom[1][4];
          var copyGrade = dtFrom[i - 1][2];
          var copyClass = dtFrom[i - 1][3];
          sheetTo.getRange(row + 1, 4).setValue(copyClientName + " " + copyGrade + "-" + copyClass); //Post to "Project Name" in This Style
          
          //Updating as "Regular"
          sheetTo.getRange(row + 1, 5).setValue("Regular");

          //Transcription of "No. Student"
          var copyNoOfStudents = dtFrom[i - 1][4];
          sheetTo.getRange(row + 1, 6).setValue(copyNoOfStudents);

          //Transcription of "Lesson Time(PHT)"
          var copyLessonTimePHT = dtFrom[i - 1][7];
          sheetTo.getRange(row + 1, 8).setValue(copyLessonTimePHT);

          //Updating "Status"
          sheetTo.getRange(row + 1, 11).setValue("Feasibility Confirmed");

          //Transcription of "Ph Ops Team Checker"
          var copyPhPic = dtFrom[3][4];
          sheetTo.getRange(row + 1, 12).setValue(copyPhPic);

          //Transcription of "Platform"
          var copyPlatform = dtFrom[4][4];
          sheetTo.getRange(row + 1, 13).setValue(copyPlatform);

          //Transcription of "Material"
          var copyMaterial = dtFrom[i - 1][8];
          sheetTo.getRange(row + 1, 16).setValue(copyMaterial);

          //Transcription of "Customize Pattern"
          var copyCustomize = dtFrom[i - 1][9];
          sheetTo.getRange(row + 1, 18).setValue(copyCustomize);

          //Transcription of "JP Sales Team PP"
          var copyJpnPic = dtFrom[2][4];
          sheetTo.getRange(row + 1, 23).setValue(copyJpnPic);

          //Updating "In Overview Tab?"
          sheetFrom.getRange(i, 12).setValue("1. Yes");

          break;
        }
      }
    }
  }
}


////A列最終行取得
//function getLastRowB(){
//　var ss = SpreadsheetApp.getActiveSpreadsheet();
//　var sheet = ss.getActiveSheet();
//　var last_row = sheet.getLastRow();
//
//  Logger.log(last_row)
//
//　for(var i = last_row; i >= 1; i--){
//　　if(sheet.getRange(i, 2).getValue() != ''){
//　　　var lastNo = i
//　　　break;
//　　}
//　}
//  Logger.log(lastNo)
//}

////A列最終行取得
//function findRow(){
//　var ss = SpreadsheetApp.getActiveSpreadsheet();
//　var sheet = ss.getActiveSheet();
//
//  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
//
//  for(var i=2; i<dat.length;i++){
//    if(dat[i][3-1] == "" ){
//      Browser.msgBox(i-1);
//      break;
//    }
//  }
//}
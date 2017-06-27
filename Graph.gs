function createChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var graphSheet = ss.getSheetByName("Graph");
  var rawSheet = ss.getSheetByName("Number");
  
  var range2017 = rawSheet.getRange(4, 7, 54, 1);
  var range2018 = rawSheet.getRange(58, 7, 53, 1);
  var range2019 = rawSheet.getRange(111, 7, 54, 1);
  var range2020 = rawSheet.getRange(165, 7, 53, 1);
  var range2021 = rawSheet.getRange(218, 7, 53, 1);
  var range2022 = rawSheet.getRange(271, 7, 53, 1);
  
  var chart = graphSheet.newChart()
              .setPosition(1,1,0,0)
              .addRange(range2017)
              .addRange(range2018)
              .addRange(range2019)
              .addRange(range2020)
              .addRange(range2021)
              .addRange(range2022)  
              .asLineChart()
              .setOption('title', 'Studying Students')
              .build();
  graphSheet.insertChart(chart);
}
var sheet = SpreadsheetApp.getActive().getActiveSheet();

function start(){
  lastlow = sheet.getLastRow();
  for(var r=2; r<=lastlow; r++){
    var range = 'D'+r+':O'+r;
    var num = sheet.getRange(r,1).getValue();
    var name = sheet.getRange(r,2).getValue();
    var title = `No.${num} ${name}`;
    var filename = `chart_${('000'+num).slice(-3)}_${name.replace(/ /,'')}.png`;
    console.log(filename);
    
    getGraph(range, title); //グラフを生成
    saveChart(filename);    //グラフを画像で保存
  }
}

function saveChart(filename) {
  var charts  = sheet.getCharts();
  var imageBlob = charts[0].getBlob().getAs('image/png').setName(filename);
  var folder = DriveApp.getFolderById('1nKM89vCEnVxcN8A2eqTB7Ijx5qKWk1rC');
  folder.createFile(imageBlob);
}


function getGraph(range, title) {
  var spreadsheet = SpreadsheetApp.getActive();
  var charts = sheet.getCharts();
  if(charts!=''){
    sheet.removeChart(charts[charts.length - 1]);
  }

  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('D1:O1'))
  .addRange(spreadsheet.getRange(range))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
  .setTransposeRowsAndColumns(true)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('title', title)
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.fontSize', 14)
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.fontSize', 14)
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setPosition(2, 16, 36, 16)
  .build();
  sheet.insertChart(chart);
};

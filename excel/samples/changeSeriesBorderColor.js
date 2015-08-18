var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.series.getItemAt(0).lineFormat.color = "#FF0000";
ctx.executeAsync().then();
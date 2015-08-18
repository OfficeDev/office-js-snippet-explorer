var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.setData("Sheet1!A1:B4", Excel.ChartSeriesBy.columns);
ctx.executeAsync().then();

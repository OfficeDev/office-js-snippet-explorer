var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.dataLabels.position = Excel.ChartDataLabelPosition.top;
ctx.executeAsync().then();
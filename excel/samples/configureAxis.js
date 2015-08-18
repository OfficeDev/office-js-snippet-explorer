var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	

chart.axes.valueAxis.maximum = 5;
chart.axes.valueAxis.minimum = 0;
chart.axes.valueAxis.majorUnit = 1;
chart.axes.valueAxis.minorUnit = 0.2;

ctx.executeAsync().then();

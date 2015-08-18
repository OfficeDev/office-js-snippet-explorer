var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	

chart.title.format.font.bold = true; 
chart.title.format.font.color = "#FF0000";

ctx.executeAsync().then();

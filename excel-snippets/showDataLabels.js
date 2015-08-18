var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	

chart.datalabels.visible = true;

ctx.executeAsync().then(function () {
		console.log("Datalabels Shown");
});
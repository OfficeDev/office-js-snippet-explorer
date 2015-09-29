
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);
	chart.datalabels.visible = true;
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
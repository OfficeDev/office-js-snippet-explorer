
Excel.run(function (ctx) {
	var range = ctx.workbook.names.getItem("MyChartData").getRange();
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.add(Excel.ChartType.pie, range, Excel.ChartSeriesBy.auto);
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
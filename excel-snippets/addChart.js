
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getItem("Sheet1").charts.add("ColumnClustered", "Sheet1!A1:D5", Excel.ChartSeriesBy.auto);
	return ctx.sync();	
});
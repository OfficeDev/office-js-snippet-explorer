
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet.charts.getItemAt(0);	
	chart.series.getItemAt(0).lineFormat.color = "#FF0000";
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
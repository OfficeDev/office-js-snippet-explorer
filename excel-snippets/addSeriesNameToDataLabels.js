
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0);	
	chart.dataLabels.showSeriesName = true;
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
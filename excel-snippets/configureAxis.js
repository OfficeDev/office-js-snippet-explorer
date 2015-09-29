
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0);	
	
	chart.axes.valueAxis.maximum = 5;
	chart.axes.valueAxis.minimum = 0;
	chart.axes.valueAxis.majorUnit = 1;
	chart.axes.valueAxis.minorUnit = 0.2;

	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
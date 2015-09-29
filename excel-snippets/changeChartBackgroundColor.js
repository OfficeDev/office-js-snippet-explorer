
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0);	
	chart.format.fill.setSolidColor("#FF0000");	
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
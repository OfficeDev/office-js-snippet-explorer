
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0);	
	chart.top = 200;
	chart.left = 200;
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});

Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0);	
	chart.axes.valueAxis.title.text = "Category";
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
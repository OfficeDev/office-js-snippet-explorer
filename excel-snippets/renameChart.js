
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0).name = "Chart1";
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
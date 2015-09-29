
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0).delete();
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
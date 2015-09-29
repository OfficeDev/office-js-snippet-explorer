
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getActiveWorksheet().delete();
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
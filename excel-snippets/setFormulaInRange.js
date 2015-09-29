
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C3").formulas = "=RAND()*17";
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
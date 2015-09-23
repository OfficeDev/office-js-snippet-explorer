
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getActiveWorksheet().getRange("A1").numberFormat = "d-mmm";
	return ctx.sync();
});
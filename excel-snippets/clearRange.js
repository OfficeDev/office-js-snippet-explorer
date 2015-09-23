
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C1").clear(Excel.ClearApplyTo.contents);	
	return ctx.sync()
});
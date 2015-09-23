
Excel.run(function (ctx) {
	ctx.workbook.names.getItem("myData").getRange().select();
	return ctx.sync();	
});
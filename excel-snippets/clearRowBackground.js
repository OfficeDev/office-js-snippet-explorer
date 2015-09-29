
Excel.run(function (ctx) {
	ctx.workbook.tables.getItem("Table1").getDataBodyRange().clear(Excel.ClearApplyTo.formats);
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
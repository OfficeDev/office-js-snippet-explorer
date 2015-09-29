
Excel.run(function (ctx) {
	ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3").values = 7;
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
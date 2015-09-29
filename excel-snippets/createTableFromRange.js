
Excel.run(function (ctx) {
	ctx.workbook.tables.add('Sheet1!A1:E7', true);
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
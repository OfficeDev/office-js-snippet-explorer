
Excel.run(function (ctx) {
	var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet().load("name");
	return ctx.sync().then(function () {
		console.log(activeWorksheet.name);
	});
}).catch(function (error) {
	console.log(error);
});

Excel.run(function (ctx) {
	var title = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0).title.load("text");
	return ctx.sync().then(function() {
		console.log(title.text);
	});
}).catch(function (error) {
	console.log(error);
});

Excel.run(function (ctx) {
	var charts = ctx.workbook.worksheets.getActiveWorksheet().charts.load("name");
	return ctx.sync().then(function () {
		for (var i = 0; i < charts.items.length; i++) {
			console.log(charts.items[i].name);
		}
		console.log("done");
	});
}).catch(function (error) {
	console.log(error);
});
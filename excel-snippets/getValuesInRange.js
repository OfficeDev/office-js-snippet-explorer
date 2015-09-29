
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C3").load("values");
	return ctx.sync().then(function () {
		for (var i = 0; i < range.values.length; i++) {
			for (var j = 0; j < range.values[i].length; j++) {
				console.log(range.values[i][j]);
			}
		}
		console.log("done");
	});
}).catch(function (error) {
	console.log(error);
});
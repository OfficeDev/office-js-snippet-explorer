
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C3").load("formulas");
	return ctx.sync().then(function() {
		for (var i = 0; i < range.formulas.length; i++) {
			for (var j = 0; j < range.formulas[i].length; j++) {
				console.log(range.formulas[i][j]);
			}
		}
		console.log("done");	
	});
}).catch(function (error) {
	console.log(error);
});
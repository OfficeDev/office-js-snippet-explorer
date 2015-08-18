var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3").load();
ctx.executeAsync().then(function() {
	for (var i = 0; i < range.formulas.length; i++) {
		for (var j = 0; j < range.formulas[i].length; j++) {
			console.log(range.formulas[i][j]);
		}
	}
	console.log("done");
});
var ctx = new Excel.RequestContext();
var worksheets = ctx.workbook.worksheets.load();
ctx.executeAsync().then(function() {
	for (var i = 0; i < worksheets.items.length; i++) {
		console.log(worksheets.items[i].name);
	}
	console.log("done");
});
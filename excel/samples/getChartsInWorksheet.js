var ctx = new Excel.RequestContext();
var charts = ctx.workbook.worksheets.getItem("Sheet1").charts.load();
ctx.executeAsync().then(function () {
	for (var i = 0; i < charts.items.length; i++) {
		console.log(charts.items[i].name);
	}
	console.log("done");
});
var ctx = new Excel.RequestContext();
var tableRows = ctx.workbook.tables.getItemAt(0).rows.load();
ctx.executeAsync().then(function () {
	for (var i = 0; i < tableRows.items.length; i++) {
		console.log(tableRows.items[i].values);
	}
	console.log("done");
});
var ctx = new Excel.RequestContext();
var tables = ctx.workbook.tables.load();
ctx.executeAsync().then(function () {
	for (var i = 0; i < tables.count; i++)
	{
		console.log(tables.items[i].name);
	}
	console.log("done");
});

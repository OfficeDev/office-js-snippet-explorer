
Excel.run(function (ctx) {
	var tables = ctx.workbook.tables.load("name");
	return ctx.sync().then(function() {
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
		console.log("done");
	});
}).catch(function (error) {
	console.log(error);
});
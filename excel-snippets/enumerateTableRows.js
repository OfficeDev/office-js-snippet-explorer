
Excel.run(function (ctx) {
	var tableRows = ctx.workbook.tables.getItemAt(0).rows.load("values");
	return ctx.sync().then(function () {
		for (var i = 0; i < tableRows.items.length; i++) {
			console.log(tableRows.items[i].values);
		}
		console.log("done");		
	});
}).catch(function (error) {
	console.log(error);
});
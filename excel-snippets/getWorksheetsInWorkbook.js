
Excel.run(function (ctx) {
	var worksheets = ctx.workbook.worksheets.load("name");
	return ctx.sync().then(function () {
		for (var i = 0; i < worksheets.items.length; i++) {
			console.log(worksheets.items[i].name);
		}
		console.log("done");
	});	
}).catch(function (error) {
	console.log(error);
});
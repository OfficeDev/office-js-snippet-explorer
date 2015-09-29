
Excel.run(function (ctx) {
	var selectedRange = ctx.workbook.getSelectedRange().load();
	return ctx.sync().then(function() {
		console.log(selectedRange.address);
	});
}).catch(function (error) {
	console.log(error);
});

Excel.run(function (ctx) {
	var range = ctx.workbook.tables.getItem('Table1').rows.getItemAt(1).getRange();
	range.format.fill.color = "#00AA00";
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});

Excel.run(function (ctx) {
	ctx.workbook.tables.getItem('Table1').rows.getItemAt(3).delete();
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
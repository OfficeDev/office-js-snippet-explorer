
Excel.run(function (ctx) {
	ctx.workbook.tables.getItem('Table1').rows.add(3, [[1,2,3,4,5]]);
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
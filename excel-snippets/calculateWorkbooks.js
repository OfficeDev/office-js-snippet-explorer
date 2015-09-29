
Excel.run(function (ctx) {
	ctx.workbook.application.calculate(Excel.CalculationType.full);
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
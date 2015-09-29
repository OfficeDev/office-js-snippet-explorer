
Excel.run(function (ctx) {
	var application = ctx.workbook.application.load("calculationMode");
	return ctx.sync().then(function () {
		console.log(application.calculationMode);			
	});	
}).catch(function (error) {
	console.log(error);
});
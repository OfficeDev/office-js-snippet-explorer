
Excel.run(function (ctx) {
	ctx.workbook.worksheets.add("Sheet" + Math.floor(Math.random()*100000).toString());
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
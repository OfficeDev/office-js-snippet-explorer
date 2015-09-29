
Excel.run(function (ctx) {
	ctx.workbook.names.getItem("myData").getRange().select();
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
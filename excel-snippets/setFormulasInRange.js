
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
	range.formulas = [["=RAND()*12", "=RAND()*19"], ["=A1*.7", "=B1*.9"]];
	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});
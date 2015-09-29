
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItemAt(0);
	chart.title.format.font.bold = true; 
	chart.title.format.font.color = "#FF0000";
	return ctx.sync();		
}).catch(function (error) {
	console.log(error);
});
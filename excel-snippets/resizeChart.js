
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
	chart.height = 200;
	chart.width = 200;
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
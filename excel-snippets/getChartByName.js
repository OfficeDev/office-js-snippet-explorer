
Excel.run(function (ctx) {
	var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.getItem("Chart1").load("name");
	return ctx.sync().then(function () {
		console.log(chart.name);		
	});
}).catch(function (error) {
	console.log(error);
});
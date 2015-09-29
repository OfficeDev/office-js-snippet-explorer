
Excel.run(function (ctx) {
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");
	
	var range = sheet.getRange("A1:B3");
	range.values = [
		["", "Gender"],
		["Male", 12],
		["Female", 14]
	];
	
	var chart = sheet.charts.add("pie", range, "auto");
	
	chart.format.fill.setSolidColor("F8F8FF");
	
	chart.title.text = "Class Demographics";
	chart.title.format.font.bold = true;
	chart.title.format.font.size = 18;
	chart.title.format.font.color = "568568";
	
	chart.legend.position = "right";
	chart.legend.format.font.name = "Algerian";
	chart.legend.format.font.size = 13;
	
	chart.dataLabels.showPercentage = true;
	chart.dataLabels.format.font.size = 15;
	chart.dataLabels.format.font.color = "444444";
	
	var points = chart.series.getItemAt(0).points;
	points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
	points.getItemAt(1).format.fill.setSolidColor("D87093");
	
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
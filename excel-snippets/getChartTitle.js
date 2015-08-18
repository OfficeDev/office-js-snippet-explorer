var ctx = new Excel.RequestContext();
var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0).title.load();	
ctx.executeAsync().then(function () {
		console.log(title.text);
		console.log("done");
});
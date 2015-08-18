var ctx = new Excel.RequestContext();
var application = ctx.workbook.application.load();
ctx.executeAsync().then(function() {
	console.log(application.calculationMode);
	console.log("done");
});

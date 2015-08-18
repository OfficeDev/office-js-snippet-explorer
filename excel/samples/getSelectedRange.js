var ctx = new Excel.RequestContext();
var selectedRange = ctx.workbook.getSelectedRange().load();
ctx.executeAsync().then(function () {
	console.log(selectedRange.address);
	console.log("done");
});
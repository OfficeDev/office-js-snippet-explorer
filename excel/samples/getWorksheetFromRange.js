var ctx = new Excel.RequestContext();
var selectedRangeWorksheet = ctx.workbook.getSelectedRange().worksheet.load();
ctx.executeAsync().then(function () {
    console.log(selectedRangeWorksheet.name);
    console.log("done");
}, function (error) {
    console.log("An error occurred: " + error.errorCode + ":" + error.errorMessage);
});
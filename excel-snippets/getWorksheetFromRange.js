
Excel.run(function (ctx) {
    var selectedRangeWorksheet = ctx.workbook.getSelectedRange().worksheet.load("name");
    return ctx.sync().then(function () {
        console.log(selectedRangeWorksheet.name);
    });
}).catch(function (error) {
	console.log(error);
});
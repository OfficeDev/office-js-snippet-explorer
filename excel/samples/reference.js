var ctx = new Excel.RequestContext();
var originalRange = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C4");
ctx.references.add(originalRange);
ctx.executeAsync()
.then(function () {
    originalRange.insert();
    originalRange.format.fill.color = "Red"; // The A5:C8(originally A1:C4 we keep reference to) will now be in red.
    ctx.references.remove(originalRange);
    console.log("The Range we keep reference to is highlighted in Red");
})
.then(ctx.executeAsync)
.catch(function (error) {
    console.log(JSON.stringify(error));
});
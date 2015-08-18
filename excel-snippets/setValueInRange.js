var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3");
range.values = 7;
ctx.executeAsync().then();
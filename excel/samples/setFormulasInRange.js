var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
range.formulas = [["=RAND()*12", "=RAND()*19"], ["=A1*.7", "=B1*.9"]];
ctx.executeAsync().then();
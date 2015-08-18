var ctx = new Excel.RequestContext();
var tableRows = ctx.workbook.tables.getItem('Table1').rows;
tableRows.add(3, [[1,2,3,4,5]]);
ctx.executeAsync().then();
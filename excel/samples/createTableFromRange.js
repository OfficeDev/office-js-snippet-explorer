var ctx = new Excel.RequestContext();
ctx.workbook.tables.add('Sheet1!A1:E7', true);
ctx.executeAsync().then();

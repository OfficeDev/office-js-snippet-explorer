var ctx = new Excel.RequestContext();
ctx.workbook.tables.getItem('Table1').rows.getItemAt(3).delete();
ctx.executeAsync().then();
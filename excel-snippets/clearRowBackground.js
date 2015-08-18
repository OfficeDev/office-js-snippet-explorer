var ctx = new Excel.RequestContext();
ctx.workbook.tables.getItem("Table1").getDataBodyRange().clear(Excel.ClearApplyTo.formats);
ctx.executeAsync().then();
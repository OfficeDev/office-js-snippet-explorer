var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getActiveWorksheet().delete();
ctx.executeAsync().then();
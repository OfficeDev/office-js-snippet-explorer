var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getActiveWorksheet().getRange("A1").numberFormat = "d-mmm";
ctx.executeAsync().then();
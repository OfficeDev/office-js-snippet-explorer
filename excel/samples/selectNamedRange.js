var ctx = new Excel.RequestContext();
ctx.workbook.names.getItem("myData").getRange().select();
ctx.executeAsync().then();
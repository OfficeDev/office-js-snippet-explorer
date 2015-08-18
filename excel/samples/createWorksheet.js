var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.add("Sheet" + Math.floor(Math.random()*100000).toString());
ctx.executeAsync().then();
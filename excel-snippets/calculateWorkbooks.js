var ctx = new Excel.RequestContext();
ctx.workbook.application.calculate(Excel.CalculationType.full);
ctx.executeAsync().then();
var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0).name = "Chart1";
ctx.executeAsync().then();

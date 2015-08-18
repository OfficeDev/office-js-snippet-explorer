var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0).delete();	
ctx.executeAsync().then();
var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0).title.visible = false; 
ctx.executeAsync().then();
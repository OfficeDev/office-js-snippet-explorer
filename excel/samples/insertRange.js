var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3").insert("right");
ctx.executeAsync().then();
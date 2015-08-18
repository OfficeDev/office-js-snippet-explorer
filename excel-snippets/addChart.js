var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.add("ColumnClustered", "Sheet1!A1:D5", Excel.ChartSeriesBy.auto);
ctx.executeAsync().then();

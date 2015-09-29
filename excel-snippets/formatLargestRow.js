
Excel.run(function (ctx) {
	var rows = ctx.workbook.tables.getItem("Table1").rows.load("values");
	return ctx.sync()
		.then(function () {
			var largestRow = 0;
			var largestValue = 0;
			
			for (var i = 0; i < rows.items.length; i++){
				if (rows.items[i].values[0][1] > largestValue){
					largestRow = i;
					largestValue = rows.items[i].values[0][1];
				}
			}
			
			var largestRowRng = rows.getItemAt(largestRow).getRange();
			largestRowRng.format.fill.color = "#ff0000";
			
		})
		.then(ctx.sync);	
}).catch(function (error) {
	console.log(error);
});
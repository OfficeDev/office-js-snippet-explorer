
Excel.run(function(ctx) {
	var range = ctx.workbook.getSelectedRange().load("values");
	return ctx.sync()
		.then(function() {
			var vals = range.values;
			for (var i = 0; i < vals.length; i++){
				for (var j = 0; j < vals[i].length; j++){
					vals[i][j] = vals[i][j].toUpperCase();
				}
			}
			range.values = vals;
		})
		.then(ctx.sync);	
}).catch(function (error) {
	console.log(error);
});
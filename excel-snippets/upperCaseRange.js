var sheetName = "Sheet1";
var rangeAddress = "E1:E5";

var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).load();

ctx.references.add(range);
ctx.executeAsync().then(function () {
	var vals = range.values;
	for(var i=0; i<vals.length;i++){
		for(var j=0;j<vals[i].length;j++){
			vals[i][j] = vals[i][j].toUpperCase();
		}
	}
	range.values = vals;
	ctx.executeAsync().then(function()
	{

		ctx.references.remove(range);
		ctx.executeAsync().then(function()
		{
		});
	});
});

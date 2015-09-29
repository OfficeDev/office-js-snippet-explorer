
//Create Table1
Excel.run(function (ctx) {
    var table = ctx.workbook.tables.add("Sheet1!A1:C4", true).load("name");
    return ctx.sync()
        .then(function() {
            Office.context.document.bindings.addFromNamedItemAsync(table.name, Office.BindingType.Table, { id: "myBinding" }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    console.log("Action failed with error: " + asyncResult.error.message)
                } else {
                    // If succeeded, then add event handler to the table binding.
                    Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            });
        });
}).catch(function (error) {
	console.log(error);
});

// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
    Excel.run(function (ctx) {
        var fill = ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.load("color");
        return ctx.sync()
            .then(function () {
                if (fill.color == "#FFA500") { 
                    return; 
                } else {
                    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "orange";
                }
            })
            .then(ctx.sync);
    }).catch(function (error) {
        console.log(JSON.stringify(error));
    });
}
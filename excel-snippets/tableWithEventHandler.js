

    //Create Table1
    var ctx = new Excel.RequestContext();
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    ctx.executeAsync()
         .then(function () {
             console.log("Table1 Created!");

             //Create a new table binding for Table1
             Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
                 if (asyncResult.status == "failed") {
                     console.log("Action failed with error: " + asyncResult.error.message);
                 }
                 else {
                     // If succeeded, then add event handler to the table binding.
                     Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                 }
             });

         })
         .catch(function (error) {
             console.log(JSON.stringify(error));
         });



// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
    var ctx = new Excel.RequestContext();
    // highlight the table in orange to indicate data has been changed.
    
    var fill= ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill;
    fill.load("color");
    ctx.executeAsync()
        .then(function () {
            if (fill.color == "#FFA500") { return; }
            else {
                ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "orange";
                console.log("The value in this table got changed!");
            }
            
        })
        .then(ctx.executeAsync)
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}

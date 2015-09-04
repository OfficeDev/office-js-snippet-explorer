/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
//Create Table1
Excel.run(function (ctx) {
    var table = ctx.workook.tables.add("Sheet1!A1:C4", true).load("name");
    return ctx.sync()
        .then(function() {
            Office.context.document.bindings.addFromNamedItemAsync(table.name, Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " asyncResult.error.message)
                } else {
                    // If succeeded, then add event handler to the table binding.
                    Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            };
        });
}).catch(function (error) {
    console.log(JSON.stringify(error));
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
/*
OfficeJS Snippet Explorer, https://github.com/OfficeDev/office-js-snippet-explorer

Copyright (c) Microsoft Corporation
All rights reserved.

MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/
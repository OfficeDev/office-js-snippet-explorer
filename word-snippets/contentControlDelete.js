/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/

//var ctx = new Word.RequestContext();
//
//// Queue: get all of the content controls in the document.
//var contentControls = ctx.document.contentControls;
//
//// Queue: load the text property for all of content controls. 
//ctx.load(contentControls, {select:"text"});
//
//// Queue: add a reference to the content controls collection.
////ctx.references.add(contentControls);
//         
//// Run the batch of commands in the queue.
//ctx.executeAsync()
//    .then(function () {
//        
//        // Queue: Delete the first content control and do not keep its contents.
//        contentControls.items[0].delete(false);
//    
//        // Queue: remove references to the content control collection.
//      //  ctx.references.remove(contentControls);
//        
//        // Run the batch of commands in the queue.
//        return ctx.executeAsync().then(
//           function () {
//               console.log("Deleted the first content control.");
//           }
//        )
//    })
//
//    .catch(function (error) {
//        console.log(JSON.stringify(error));
//    });

// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to get the first content control in the document.
    // Create a proxy object for a content control.
    var contentControl = contentControls.getItem(0);
     
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControl === null) {
            console.log("There isn't a content control in this document.");
        } else {
            // Queue a command to delete the content control.
            contentControl.delete(false);
            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Deleted content control.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});


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
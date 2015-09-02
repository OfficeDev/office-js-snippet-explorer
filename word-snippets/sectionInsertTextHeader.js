/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
var ctx = new Word.RequestContext();

// Queue: get all of the sections in the document.
var mySections = ctx.document.sections;

// Queue: load the sections.
ctx.load(mySections, { select: 'text' });

// Queue: add a reference to the sections collection.
ctx.references.add(mySections);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
    
        // Queue: get the primary header of the first section. Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
    
        // Queue: insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
    
        // Queue: wrap the header in a content control.
        myHeader.insertContentControl();

        // Queue: remove the references to the sections.
        ctx.references.remove(mySections);
    
    })
    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Added a header to the first section.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
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
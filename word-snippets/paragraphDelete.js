/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs.
ctx.load(paragraphs, { select: "text" });

// Queue: add a reference to the paragraphs collection.
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph from the collection.
        var paragraph = paragraphs._GetItem(0);

        // Queue: delete the paragraph.
        paragraph.delete();

        // Queue: remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);

        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(
           function () {
               console.log("Deleted the paragraph.");
           }
        )
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
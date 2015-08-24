/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs and their text property.
ctx.load(paragraphs, { select: "text" });

// Queue: add a reference to the paragraphs collection
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue: load font and style information for the paragraph.
        ctx.load(paragraph, {select: 'font/size, font/name, font/color, style',
                             expand: 'font'});

        // Queue: Remove the reference to the paragraphs.
        ctx.references.removeAll();

        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(function () {

            // Show the results of the load method. Here we show the
            // property values on the paragraph object.
            var results = "<strong>Paragraph</strong><br>" +
                          "<br>Font size: " + paragraph.font.size +
                          "<br>Font name: " + paragraph.font.name +
                          "<br>Font color: " + paragraph.font.color +
                          "<br>Style: " + paragraph.style;

            console.log(results);
        });
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
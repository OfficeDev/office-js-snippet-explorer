/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
var ctx = new Word.RequestContext();

// Setup the search options.
var options = Word.SearchOptions.newObject(ctx);
options.matchCase = false

// Queue: search the document.
var searchResults = ctx.document.body.search('Video', options);

// Queue: load the results and get the font property values.
ctx.load(searchResults, { expand: 'font' });

// Queue: add a reference to the search results collection.
ctx.references.add(searchResults);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue: change the font for each found item. Select the 4th item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = '#FF0000'; //Red
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
            if (i === 3)
                searchResults.items[i].select();
        }

        // Queue: remove the reference to the search results.
        ctx.references.remove(searchResults);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log('Highlighted the search results.');
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
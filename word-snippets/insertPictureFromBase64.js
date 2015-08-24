/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);
ctx.references.add(paras);

ctx.executeAsync()
    .then(function () {
        paras.items[0].insertInlinePictureFromBase64("iVBORw0KGgoAAAANSUhEUgAAAIAAAACABAMAAAAxEHz4AAAAJFBMVEX///9GRkZGRkZGRkZGRkZGRkZGRkZGRkYBpO9/ugDyUCL/uQGm4PjWAAAACHRSTlMBCQ0RFRknMx7uViEAAAB3SURBVGje7dcxCYBQGEXhi6izYBHB0RIiiAXkzW5iAMEKFnCwguVscJd/ecM5Ab79SNHK5FqlZXeNql/XIx23awMAAAAAAAAAAAAAAAAAyBwIvzNJxeyapLZ3Naou1ykNn6sDAAAAAAAAAAAAAAAAAMgcCL9ztB/UhshWs1l/WAAAAABJRU5ErkJggg==", Word.InsertLocation.end);
        var pics = ctx.document.body.inlinePictures
        ctx.load(pics);

        ctx.executeAsync()
            .then(function () {
                console.log("Picture Count: " + pics.items.length);
                console.log("Success");
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
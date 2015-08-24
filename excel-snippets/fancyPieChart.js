/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.*/
var ctx = new Excel.RequestContext();
var sheet = ctx.workbook.worksheets.getItem("Sheet1");
 
var range = sheet.getRange("A1:B3");
range.values = [
	   ["", "Gender"],
	   ["Male", 12],
	   ["Female", 14]
];
 
var chart = sheet.charts.add("pie", range, "auto");
 
chart.format.fill.setSolidColor("F8F8FF");
 
chart.title.text = "Class Demographics";
chart.title.format.font.bold = true;
chart.title.format.font.size = 18;
chart.title.format.font.color = "568568";
 
chart.legend.position = "right";
chart.legend.format.font.name = "Algerian";
chart.legend.format.font.size = 13;
 
chart.dataLabels.showPercentage = true;
chart.dataLabels.format.font.size = 15;
chart.dataLabels.format.font.color = "444444";
 
var points = chart.series.getItemAt(0).points;
points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
points.getItemAt(1).format.fill.setSolidColor("D87093");
 
ctx.executeAsync().then();
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
var ctx = new Word.RequestContext();

// Queue: get the selected text in the document.
var objRange = ctx.document.getSelection();

// Queue: load font and style information for the range.
ctx.load(objRange, {
    select: 'font/size, font/name, font/color, style',
    expand: 'font'
});

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Show the results of the load method. Here we show the
        // property values on the range object.
        var results = "<strong>Range</strong><br>" +
                      "<br>Font size: " + objRange.font.size +
                      "<br>Font name: " + objRange.font.name +
                      "<br>Font color: " + objRange.font.color +
                      "<br>Style: " + objRange.style;
        console.log(results);

    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
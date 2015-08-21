var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: load font and style information for the document body.
ctx.load(body, {select: 'font/size, font/name, font/color, style',
                expand: 'font'});

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = "<strong>Body</strong><br>" +
                      "<br>Font size: " + body.font.size +
                      "<br>Font name: " + body.font.name +
                      "<br>Font color: " + body.font.color +
                      "<br>Style: " + body.style;

        console.log(results);
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
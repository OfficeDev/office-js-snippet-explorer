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
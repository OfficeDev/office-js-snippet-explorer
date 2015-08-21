var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs.
ctx.load(paragraphs, { select: "text" });

// Queue: add a reference to the paragraphs collection
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph.
        var paragraph = paragraphs._GetItem(0);

        // Queue: wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl();

        // Queue: remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Wrapped the first paragraph in a content control.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
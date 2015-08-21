var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs and their text.
ctx.load(paragraphs, { select: "text" });

// Queue: add a reference to the paragraphs collection
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph.
        var paragraph = paragraphs._GetItem(0);

        // Queue: Insert HTML content at the end of the first paragraph.
        paragraph.insertHtml("<strong>Inserted HTML.</strong>", Word.InsertLocation.end);

        // Queue: Remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Insert HTML content at the end of the first paragraph.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
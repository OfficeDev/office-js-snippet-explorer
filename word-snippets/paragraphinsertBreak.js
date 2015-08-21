var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs.
ctx.load(paragraphs, { select: "text",
                       expand: "paragraph"});

// Queue: add a reference to the paragraphs collection
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue: insert a page break after the first paragraph.
        paragraph.insertBreak("page", "After");

        // Queue: remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Inserted a page break after the paragraph.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
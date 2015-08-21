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

        // Queue: insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

        // Queue: remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Inserted a new paragraph at the end of the first paragraph.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
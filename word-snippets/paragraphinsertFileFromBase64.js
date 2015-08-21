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

        // Queue: insert base64 encoded .docx at the beginning of the first paragraph.
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

        // Queue: remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Insert base64 encoded content at the beginning of the first paragraph.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
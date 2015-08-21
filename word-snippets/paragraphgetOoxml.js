var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs.
ctx.load(paragraphs, { select: "text",
                       expand: "paragraph"});

// Queue: add a reference to the paragraphs collection.
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph from the collection.
        var paragraph = paragraphs.items[0];

        // Queue: get the OOXML representation of the paragraph.
        var result = paragraph.getOoxml();

        // Queue: remove the reference to the paragraphs. 
        ctx.references.remove(paragraphs);

        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(
           function () {
               console.log(result.value);
           }
        )

    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
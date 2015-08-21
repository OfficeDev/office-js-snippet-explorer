var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs.
ctx.load(paragraphs, { select: "text" });

// Queue: add a reference to the paragraphs collection.
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph from the collection.
        var paragraph = paragraphs._GetItem(0);

        // Queue: delete the paragraph.
        paragraph.delete();

        // Queue: remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);

        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(
           function () {
               console.log("Deleted the paragraph.");
           }
        )
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
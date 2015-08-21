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

        // Queue: clear the contents of the first paragraph.
        paragraphs._GetItem(0).clear();

        // Queue: remove references to the paragraphs collection.
        ctx.references.remove(paragraphs);
        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(
           function () {
               console.log("Cleared the paragraph.");
           }
        )
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
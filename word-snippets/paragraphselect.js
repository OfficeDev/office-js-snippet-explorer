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

        // Queue: get the last paragraph.
        var paragraph = paragraphs.items[paragraphs.items.length - 1];

        // Queue: select the paragraph. The Word UI will 
        // move to the selected paragraph.
        paragraph.select();

        // Queue: Remove the reference to the paragraphs collection.
        ctx.references.removeAll();

        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(function () {
            console.log("Selected the paragraph.");
        });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });  
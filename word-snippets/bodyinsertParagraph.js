var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: insert the paragraph at the end of the document body.
body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Paragraph added at the end of the document body.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

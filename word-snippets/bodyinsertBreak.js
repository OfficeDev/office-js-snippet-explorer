var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: insert a line break at the start of the document body.
body.insertBreak(Word.BreakType.line, Word.InsertLocation.start);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Added a line break at the start of the document body.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
}
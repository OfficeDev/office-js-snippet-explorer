var ctx = new Word.RequestContext();

// Queue: get the selected text in the document.
var objRange = ctx.document.getSelection();

// Queue: insert text at the end of the range.
objRange.insertText('New text inserted into the range.', Word.InsertLocation.end);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        console.log("Text added to the end of the range.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
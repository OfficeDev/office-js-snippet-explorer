var ctx = new Word.RequestContext();

// Queue: get the selected text from the document.
var objRange = ctx.document.getSelection();

// Queue: insert the paragraph after the range.
objRange.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        console.log("Paragraph added to the end of the range.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
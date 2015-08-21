var ctx = new Word.RequestContext();

// Queue: get the selected text in the document and create a range object.
var objRange = ctx.document.getSelection();

// Queue: insert a page break after the selected text.
objRange.insertBreak("page", "After");

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        console.log("Inserted a page break after the selected text.");
    })

 .catch(function (error) {
     console.log(JSON.stringify(error));
 });
var ctx = new Word.RequestContext();

// Queue: get the selected text from the document.
var objRange = ctx.document.getSelection();

// Queue: insert base64 encoded .docx at the beginning of the range.
objRange.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Added base64 encoded text to the beginning of the range.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
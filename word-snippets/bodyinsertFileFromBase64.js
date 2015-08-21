var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: insert base64 encoded .docx at the beginning of the content body.
// "WWVzIQ=="
body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);



// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Added base64 encoded text to the beginning of the document body.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

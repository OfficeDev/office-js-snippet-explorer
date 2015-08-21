
// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: get the OOXML contents of the body.
var bodyHTML = body.getOoxml();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Body OOXML contents: " + bodyHTML.value);
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
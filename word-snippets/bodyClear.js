var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: clear the contents of the body.
body.clear();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Cleared the body contents.")
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
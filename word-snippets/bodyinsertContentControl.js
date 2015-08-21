var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: wrap the body in a content control.
body.insertContentControl();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Wrapped the body in a content control.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
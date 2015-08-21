var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: insert HTML in to the beginning of the body.
body.insertHtml("<strong>This is text inserted with body.insertHtml()</strong>", Word.InsertLocation.start);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("HTML added to the beginning of the document body.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
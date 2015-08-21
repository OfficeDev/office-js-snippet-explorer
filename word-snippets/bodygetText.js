var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: load the text in document body.
ctx.load(body, { select: "text" });

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Body contents: " + body.text);
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
}
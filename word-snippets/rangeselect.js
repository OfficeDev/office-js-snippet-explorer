var ctx = new Word.RequestContext();

// Queue: get the document body.
var body = ctx.document.body;

// Queue: insert HTML in to the beginning of the body.
var range = body.insertHtml("<strong>This is text inserted with body.insertHtml()</strong>", Word.InsertLocation.start);

// Queue: select the HTML that was inserted.
range.select();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        console.log("Selected the range of inserted HTML content.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
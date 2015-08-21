var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: get the HTML contents of the body.
var bodyHTML = body.getHtml();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Body HTML contents: " + bodyHTML.value);
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
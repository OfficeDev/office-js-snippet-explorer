var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: select the docment body. The Word UI will 
// move to the selected document body.
body.select();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("Selected the document body.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
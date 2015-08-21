var ctx = new Word.RequestContext();

// Queue: get the user's current selection and create a range object named objrange.
var objRange = ctx.document.getSelection();

// Queue: insert HTML in to the beginning of the range.
objRange.insertHtml("<strong>This is text inserted with range.insertHtml()</strong>", Word.InsertLocation.start);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        console.log("HTML added to the beginning of the range.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
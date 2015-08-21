var ctx = new Word.RequestContext();

// Queue: get the user's current selection and create a range object named objrange.
var objRange = ctx.document.getSelection();

// Queue: get the OOXML representation of the range.
var result = objRange.getOoxml();

// Run the batch of commands in the queue. 
ctx.executeAsync()
    .then(function () {

        console.log("The OOXML read from the document was: \r\n\r\n" + result.value);
    })
    .catch(function (error) {
        console.log("ERROR: " + JSON.stringify(error));
    });
var ctx = new Word.RequestContext();

// Queue: get the user's current selection and create a range object named objrange.
var objRange = ctx.document.getSelection();

// Queue: get the user's current selection and create a range object named objrange.
var result = objRange.getHtml();

// Run the batch of commands in the queue. 
ctx.executeAsync()
    .then(function () {

        console.log("The HTML read from the document was: \r\n\r\n" + result.value);
    })
    .catch(function (error) {
        console.log("ERROR: " + JSON.stringify(error));
    });
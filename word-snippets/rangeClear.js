var ctx = new Word.RequestContext();

// Queue: get the user's current selection and create a range object named objrange.
// Queue: clear objrange.
var objrange = ctx.document.getSelection();
objrange.clear();

// Run the set of commands in the queue. In this case, we are clearing the range. 
ctx.executeAsync()
    .then(function () {
        console.log("Done");
    })
    .catch(function (error) {
        console.log("ERROR: " + JSON.stringify(error));
    });
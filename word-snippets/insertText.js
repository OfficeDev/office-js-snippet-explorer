var ctx = new Word.RequestContext();
var range = ctx.document.getSelection();

var textSample =
    "Hello, world! This is an example of the insert text method. This is a method which allows users to insert text into a given selection. It can insert text into a relative location or it can overwrite the current selection.";

range.insertText(textSample, Word.InsertLocation.end);

ctx.executeAsync()
    .then(function () {
         console.log("Success");
     })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

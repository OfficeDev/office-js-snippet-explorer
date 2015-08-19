var ctx = new Word.RequestContext();
var range = ctx.document.getSelection();

var myContentControl = range.insertContentControl();
myContentControl.tag = "Customer-Address";
myContentControl.title = "Enter Customer Address Here:";
myContentControl.style = "Heading 1";
myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
myContentControl.cannotEdit = true;
myContentControl.appearance = "tags";

ctx.load(myContentControl);

ctx.executeAsync()
    .then(function () {
        console.log("Content control Id: " + myContentControl.id);
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

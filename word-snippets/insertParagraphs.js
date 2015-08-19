var ctx = new Word.RequestContext();

var myPar = ctx.document.body.insertParagraph("Bibliography", "end");
myPar.style = "Heading 1";

var myPar2 = ctx.document.body.insertParagraph("This is my first book.", "end");
myPar2.style = "Normal"

ctx.executeAsync()
    .then(function () {
        console.log("Success");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

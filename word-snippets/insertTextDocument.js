var ctx = new Word.RequestContext();

var textSample =
    "Hello, world! This is an example of the insert text method. This is a method which allows users to insert text at the end of the document. It also can insert text into a relative location.";

ctx.document.body.insertParagraph(textSample, Word.InsertLocation.end);

ctx.executeAsync().then(
     function () {
         console.log("Success");
     },
     function (result) {
         console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
         console.log(result.traceMessages);
     }
);

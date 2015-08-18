var ctx = new Word.RequestContext();

var mySections = ctx.document.sections;
ctx.load(mySections);

var paras = mySections.getItem(0).body.paragraphs;
ctx.load(paras);

ctx.executeAsync().then(
    function () {
        console.log("Number of paragraphs in section: " + paras.items.length);
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);
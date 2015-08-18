var ctx = new Word.RequestContext();
var ccs = ctx.document.contentControls.getByTag("Customer-Address");
ctx.load(ccs);
ccs.getItemAt(0).font.italic = true;

ctx.executeAsync().then(
    function () {
        var ccText = ccs.getItemAt(0).getText();
        ctx.executeAsync().then(
            function () {
                console.log("Content Control Text: " + ccText.value);
            }
         );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);

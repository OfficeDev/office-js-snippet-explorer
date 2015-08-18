var ctx = new Word.RequestContext();
var cCtrls = ctx.document.body.contentControls;
ctx.load(cCtrls);

ctx.executeAsync().then(
    function () {
        var results = new Array();
        for (var i = 0; i < cCtrls.count; i++) {
            results.push(cCtrls.getItemAt(i));
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("contentControl[" + i + "].length = " + results[i].text.length);
                }
            }
        );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);

var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);
ctx.references.add(paras);

ctx.executeAsync().then(
    function () {
        var font = paras.getItem(0).font;
        font.size = 32;
        font.bold = true;
        font.color = "#0000ff";
        font.highlightColor = "#ffff00";

        ctx.references.remove(paras);
        ctx.executeAsync().then(
            function () {
                console.log("Success");
            }
        );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);

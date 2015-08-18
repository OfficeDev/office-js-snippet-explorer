var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);

var par = paras.getItem(0);
par.lineSpacing = 36;

ctx.load(par);
var val = par.lineSpacing;

ctx.executeAsync().then(
	function () {
		console.log("Success! Setting paragraph line spacing to " + val);
	},
	function (result) {
		console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
		console.log(result.traceMessages);
	}
);

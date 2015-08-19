var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);
ctx.references.add(paras);

ctx.executeAsync()
    .then(function () {
        var par = paras.items[0];
        par.lineSpacing = 36;

        ctx.load(par);
        var val = par.lineSpacing;

        ctx.references.remove(paras);
        ctx.executeAsync()
            .then(function () {
                console.log("Success! Setting paragraph line spacing to " + val);
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

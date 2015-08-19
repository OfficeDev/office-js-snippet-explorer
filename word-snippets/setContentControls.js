var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);
ctx.references.add(paras);

ctx.executeAsync()
    .then(function () {
        for (var i = 0; i < paras.items.length; i++) {
            paras.items[i].insertContentControl();
        }

        ctx.references.remove(paras);
        ctx.executeAsync()
            .then(function () {
                console.log("Success");
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
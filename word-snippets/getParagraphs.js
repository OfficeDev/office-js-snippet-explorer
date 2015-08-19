var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras, {select: "text"});

ctx.executeAsync()
    .then(function () {
        var results = new Array();
        for (var i = 0; i < paras.items.length; i++) {
            results.push(paras.items[i].text);
        }

        ctx.executeAsync()
            .then(function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("paras[" + i + "].content  = " + results[i]);
                }
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

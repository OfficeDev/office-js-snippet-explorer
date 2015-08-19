var ctx = new Word.RequestContext();
var cCtrls = ctx.document.body.contentControls;
ctx.load(cCtrls, { select: "text" });
ctx.references.add(cCtrls);

ctx.executeAsync()
    .then(function () {
        var results = new Array();
        for (var i = 0; i < cCtrls.items.length; i++) {
            results.push(cCtrls.items[i]);
        }

        ctx.references.remove(cCtrls);
        ctx.executeAsync()
            .then(function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("contentControl[" + i + "].length = " + results[i].text.length);
                }
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

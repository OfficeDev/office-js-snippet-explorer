var ctx = new Word.RequestContext();
var ccs = ctx.document.contentControls.getByTag("Customer-Address");
ctx.load(ccs);
ctx.references.add(ccs);

ctx.executeAsync()
    .then(function () {
        ctx.references.remove(ccs);
        ctx.executeAsync().then(
            function () {
                console.log("Content Control Text: " + ccs.items[0].text);
            }
         );
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

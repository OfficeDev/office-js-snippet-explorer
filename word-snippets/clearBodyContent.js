var ctx = new Word.RequestContext();
ctx.document.body.clear();

ctx.executeAsync()
    .then(function () {
        console.log("Success");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

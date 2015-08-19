var ctx = new Word.RequestContext();

var mySections = ctx.document.sections;
ctx.load(mySections);
ctx.references.add(mySections);

ctx.executeAsync()
    .then(function () {
        var myHeader = mySections.items[0].getHeader("primary");
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        myHeader.insertContentControl();

        ctx.executeAsync()
        .then(function () {
            ctx.references.remove(mySections);
            console.log("Success");
        });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
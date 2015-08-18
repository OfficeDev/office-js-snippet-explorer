var ctx = new Word.RequestContext();

var mySections = ctx.document.sections;
ctx.load(mySections);

var myHeader = mySections.getItem(0).getHeader("primary");
myHeader.insertText("This is a header.", "end");
myHeader.insertContentControl();

ctx.executeAsync().then(
    function () {
        console.log("Success");
    }
);

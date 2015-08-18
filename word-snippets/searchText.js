var ctx = new Word.RequestContext();
var options = Word.SearchOptions.newObject(ctx);

options.matchCase = false

var results = ctx.document.body.search("Video", options);
ctx.load(results);
ctx.references.add(results);

ctx.executeAsync().then(
    function () {
        console.log("Found count: " + results.items.length);
        for (var i = 0; i < results.items.length; i++) {
            results.items[i].font.color = "#FF0000"    // Change color to Red
            results.items[i].font.highlightColor = "#FFFF00";
            results.items[i].font.bold = true;
            if (i == 3)
                results.items[i].select();
        }
        ctx.references.remove(results);
        ctx.executeAsync().then(
            function () {
                console.log("Deleted");
            }
        );
    }
);

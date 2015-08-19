var ctx = new Word.RequestContext();
var results = ctx.document.body.search("Video");

ctx.load(results, { expand: "font" });
ctx.references.add(results);

ctx.executeAsync()
    .then(function () {
        console.log("Found count: " + results.items.length);

        for (var i = 0; i < results.items.length; i++) {
            results.items[i].font.color = "#FF0000"; //Red
            results.items[i].font.highlightColor = "#FFFF00"; //Yellow
            results.items[i].font.bold = true;
            if (i == 3)
                results.items[i].select();
        }

        ctx.references.remove(results);
        ctx.executeAsync()
            .then(function () {
                console.log("Completed!");
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });

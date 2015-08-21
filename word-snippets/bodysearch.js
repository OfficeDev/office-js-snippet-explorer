var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: setup the search options.
var options = Word.SearchOptions.newObject(ctx);
options.matchCase = false

// Queue: search the document.
var searchResults = ctx.document.body.search("Video", options);

// Queue: load the results.
ctx.load(searchResults, {select:"text, font/color", 
                         expand:"font"});

// Queue: add a reference to the results.
ctx.references.add(searchResults);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        var results = "Found count: " + searchResults.items.length + 
                      "<br>We highlighted the results and selected the 4th item.";

        // Queue: Change the font for each found item. Select the 4th item.
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = "#FF0000"    // Change color to Red
          searchResults.items[i].font.highlightColor = "#FFFF00";
          searchResults.items[i].font.bold = true;
          if (i == 3)
            searchResults.items[i].select();
        }

        // Queue: remove the reference to the search results.
        ctx.references.remove(searchResults);

        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(function () {
            console.log(results);
        });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
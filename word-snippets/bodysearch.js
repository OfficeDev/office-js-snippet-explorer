// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Setup the search options.
    var options = Word.SearchOptions.newObject(context);
    options.matchCase = false

    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', options);

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length + 
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item. 
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log(results);
        });  
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

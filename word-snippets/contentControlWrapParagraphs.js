// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a command to load the style property for all of the paragraphs. 
    context.load(paragraphs, 'style');
     
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (paragraphs.items.length === 0) {
            console.log('No paragraphs found.');
        }
        else {
            for (var i = 0; i < paragraphs.items.length; i++) {
                // Queue a command to wrap each paragraph in a content control.
                paragraphs.items[i].insertContentControl();
            }
        
            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Wrapped all paragraphs in a content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
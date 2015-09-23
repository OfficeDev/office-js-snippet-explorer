// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');
        
        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we 
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong><br>" +
                          "<br>Font size: " + paragraph.font.size +
                          "<br>Font name: " + paragraph.font.name +
                          "<br>Font color: " + paragraph.font.color +
                          "<br>Style: " + paragraph.style;

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
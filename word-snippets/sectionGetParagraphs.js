// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the first section.
    context.load(mySections, {select: 'body/style', top: 1});
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object for the paragraphs collection in the first section.
        var paragraphs = mySections.items[0].body.paragraphs;
        
        // Queue a command to load the paragraphs and their text property.
        context.load(paragraphs, 'text');
                              
        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Number of paragraphs in section: " + paragraphs.items.length);
        });                    
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
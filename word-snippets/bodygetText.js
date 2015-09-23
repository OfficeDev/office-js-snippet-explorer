// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load the text in document body.
    context.load(body, 'text');
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});

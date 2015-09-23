// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to select the document body. The Word UI will 
    // move to the selected document body.
    body.select();
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

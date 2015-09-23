// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.thick;
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
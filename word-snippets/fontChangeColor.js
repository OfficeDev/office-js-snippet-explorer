// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the font color of the current selection.
    selection.font.color = 'blue'; 
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font color of the selection has been changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to insert a paragraph at the end of the body and
    // create a proxy object for the paragraph. Set the style 
    // of the proxy paragraph object.
    var myPar = context.document.body.insertParagraph('Bibliography', 'end');
    myPar.style = 'Heading 1';
    
    // Queue a command to insert a paragraph at the end of the body and
    // create a proxy object for the paragraph. Set the style 
    // of the proxy paragraph object.
    var myPar2 = context.document.body.insertParagraph('This is my first book.', 'end');
    myPar2.style = 'Normal'
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added a bibliography section to the end of the body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
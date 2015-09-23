// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var htmlText =
    '<h1><strong>Insert Html</strong></h1>' +
    '<h2><em>Office Extensibility Platform</em></h2>' +
    '<p>This is an example of how the InsertHtml method works.</p>' +
    '<table>' +
        '<tr><td>Check</td><td>out</td></tr>' +
        '<tr><td>this</td><td>table</td></tr>' +
    '</table>';
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert HTML at the end of the selection.
    range.insertHtml(htmlText, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the HTML at the end of the selection.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
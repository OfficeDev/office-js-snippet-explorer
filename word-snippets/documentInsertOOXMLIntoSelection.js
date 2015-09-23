// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var ooxmlText = "<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'>" + 
    "<w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/>" + 
    "</w:rPr><w:t>Hello world (this should be bold, red, size 14).</w:t></w:r></w:p>";
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert OOXML at the end of the selection.
    range.insertOoxml(ooxmlText, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the OOXML at the end of the selection.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
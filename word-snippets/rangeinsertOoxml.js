// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'> " +
                      "<w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/>" +                           
                      "<w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this should be bold," +
                      "red, size 14).</w:t></w:r></w:p>", Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

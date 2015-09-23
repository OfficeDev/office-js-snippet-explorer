// Run a batch operation against the Word object model.
Word.run(function (ctx) {
    
    // Create a proxy object for the document body.
    var body = ctx.document.body;
    
    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


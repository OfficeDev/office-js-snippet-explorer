var ctx = new Word.RequestContext();

// Queue: get a handle on the document body.
var body = ctx.document.body;

// Queue: insert OOXML in to the beginning of the body.
body.insertOoxml("<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this should be bold, red, size 14).</w:t></w:r></w:p>", Word.InsertLocation.start);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        console.log("OOXML added to the beginning of the document body.");
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
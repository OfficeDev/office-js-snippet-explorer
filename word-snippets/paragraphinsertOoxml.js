var ctx = new Word.RequestContext();

// Queue: get all of the paragraphs in the document.
var paragraphs = ctx.document.body.paragraphs;

// Queue: load the paragraphs.
ctx.load(paragraphs, { select: "text" });

// Queue: add a reference to the paragraphs collection
ctx.references.add(paragraphs);

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {

        // Queue: get the first paragraph.
        var paragraph = paragraphs._GetItem(0);

        // Queue: insert Ooxml content into the first paragraph.
        var ooxmlContent = "<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this should be bold, red, size 14).</w:t></w:r></w:p>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);

        // Queue: Remove the reference to the paragraphs.
        ctx.references.remove(paragraphs);
    })

    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log("Insert OOXML at the end of the first paragraph.");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
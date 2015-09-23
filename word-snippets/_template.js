var ctx = new Word.RequestContext();

// Run the batch of commands in the queue.
ctx.executeAsync()
    .then(function () {
        //Act on objects.
    
        // Use this if you nest actions.
        // Run the batch of commands in the queue.
        return ctx.executeAsync().then(function () {
            // Do something here.
        });
    })
    // Run the batch of commands in the queue.
    .then(ctx.executeAsync)
    .then(function () {
        console.log('Success')
    })

    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
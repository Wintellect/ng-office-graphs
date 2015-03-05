
Office.initialize = function () {
    console.log(">>> Office.initialize()");

    Office.context.document.bindings.addFromPromptAsync('matrix', function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult);

            asyncResult.value.getDataAsync(function(data) {
                console.log(data.value);

                drawObject.draw(data.value);

            });
        }
    });


};


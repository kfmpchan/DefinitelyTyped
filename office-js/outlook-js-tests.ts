/*
outlook-js-tests
Copyright (c) Microsoft Corporation
*/

/// <reference path="office-js.d.ts" />
function test_outlook() {
    try {
        const methodName = "getCallbackTokenAsync";
        var asyncContext: string;
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true, asyncContext: methodName}, asyncResultCallback);
    }
    catch (error) {
        console.log('Error: ' + JSON.stringify(error));
    }

    try {
        const methodName = "item.SaveAsync";        
        var composeItem = Office.cast.item.toMessageCompose(Office.context.mailbox.item);
        composeItem.saveAsync({ asyncContext: methodName }, asyncResultCallback);
    }
    catch (error) {
        console.log('Error: ' + JSON.stringify(error));
    }    
}

function asyncResultCallback(result: Office.AsyncResult): void {
    if (result.status != Office.AsyncResultStatus.Succeeded) {
        console.log(result.asyncContext + " returned an error " + JSON.stringify(result.error));
    }
}

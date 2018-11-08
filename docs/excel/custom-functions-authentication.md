---
title: Authentication using Custom Functions
description: Learn how to authenticate users via Excel Custom Functions.
ms.date: 11/6/2018
---

# Authentication

You may wish to verify that a user is authenticated before allowing them access to your custom functions. Implementing authentication for custom functions does not differ significantly in process from most Office add-ins, but there are some specific APIs used by custom functions that you should use.
  
## AsyncStorage

The `AsyncStorage` object is common and accessible to both custom functions and UI elements of your add-in such as the task pane. Because of this commonality, you can use it to store information that needs to pass back and forth between these parts of your add in.

For example, if a user enters their credentials through a UI element like the task pane, you should store the resulting tokens in `AsyncStorage` if your custom function makes use of them.  

`AsyncStorage` also offers a sandboxed environment on a user's device and cannot be accessed by other add-ins.  Note that there are some locations which should not be used to store data if you are using custom functions:  

- `localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in [`localStorage`](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage).

- `Office.context.document.settings`:  `Office.context.document.settings` is not secure and information can be extracted by anyone that has access to the document either directly, through your add-in or through another add-in.

## Dialog API

If your function checks `AsyncStorage` and does not find an access token in the process of authenticating, you should use the [`DisplayWebDialog`](https://docs.microsoft.com/en-us/javascript/api/office-runtime/officeruntime.displaywebdialogoptions?view=office-js) API to prompt the user to enter their credentials.  
  
As mentioned previously, when a user enters their credentials via the dialog box, tokens should be stored using `AsyncStorage`.  
  
The following code sample shows how you can use the `displayWebDialog` to display a dialog box for that purpose.

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"
  
function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {
  
      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {
  
        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else { 
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }
  
  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
        let timeout = 5;
        let count = 0;
        var intervalId = setInterval(function () {
          count++;
          if(_cachedToken) {
            resolve(_cachedToken);
            clearInterval(intervalId);
          }
          if(count >= timeout) {
            reject("Timeout while waiting for token");
            clearInterval(intervalId);
          }
        }, 1000);
      } else {
        _dialogOpen = true;
        OfficeRuntime.displayWebDialogOptions(url, {
          height: '50%',
          width: '50%',
          onMessage: function (message, dialog) {
            _cachedToken = message;
            resolve(message);
            dialog.closeDialog();
            return;
          },
          onRuntimeError: function(error, dialog) {
            reject(error);
          },
        }).catch(function (e) {
          reject(e);
        });
      }
    });
  }
}
```

## Single sign-on

Single sign-on is not presently available for custom functions. For guidance for Office add-ins which do not use custom functions, see [Enable single sign-on for Office Add-ins (preview)](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins).

## See also

* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions overview](custom-functions-overview.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](excel-tutorial-custom-functions.md)

/*
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

var dialog;
var sendEvent;

function validateRecipients(event) {

  if (Office.context.requirements.isSetSupported("MailBox", "1.3") && Office.context.requirements.isSetSupported("DialogApi", "1.2") )
  {
    sendEvent = event;

    let item = Office.context.mailbox.item;
  
    if (item.itemType === "message")
    {
      let toRecipientPromise = getToRecipients(item);
      let ccRecipientPromise = getCcRecipients(item);
      let bccRecipientPromise = getBccRecipients(item);
    
      Promise.all([toRecipientPromise, ccRecipientPromise, bccRecipientPromise]).then((promises) => {  
        const combinedRecipients = [...promises[0],...promises[1],...promises[2]];
        let externalRecipients = getExternalRecipients(combinedRecipients);
        handleDialog(externalRecipients, event);
      });
    }
    else if (item.itemType === "appointment")
    {
      let requiredRecipientPromise = getRequiredAttendees(item);
      let optionalRecipientPromise = getOptionalAttendees(item);
    
      Promise.all([requiredRecipientPromise, optionalRecipientPromise]).then((promises) => {
        const combinedRecipients = [...promises[0],...promises[1]];
        let externalRecipients = getExternalRecipients(combinedRecipients);
        handleDialog(externalRecipients, event);
      });
    }
  }
}

function handleDialog(externaRecipients, event) {
  if (externaRecipients.length > 0) {
    
    const topRecipients = externaRecipients.slice(0,3).map((item) => {return item.emailAddress;}); //TODO configurable slice length
    const recipientString = JSON.stringify(topRecipients);
    window.localStorage.setItem('recipients', recipientString);

    Office.context.ui.displayDialogAsync(window.location.origin + '/validate.html', { height: 18, width: 30, promptBeforeOpen: false, displayInIframe: true}, 
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    });
  } 
  else {
    event.completed({ allowEvent: true });
  }
}

function getExternalRecipients(recipients) {
  let externalRecipients = [];
  for (let i = 0; i < recipients.length; i++) {
    if (recipients[i].recipientType === "externalUser") {
      externalRecipients.push(recipients[i]);
    }
  }

  return externalRecipients;
}

function getRequiredAttendees(item) {
    return new Office.Promise(function (resolve, reject) {
      try {
        item.requiredAttendees.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject(error);
      }
  })
}

function getOptionalAttendees(item) {
    return new Office.Promise(function (resolve, reject) {
      try {
        item.optionalAttendees.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject(error);
      }
  })
}

function getToRecipients(item) {
    return new Office.Promise(function (resolve, reject) {
      try {
        item.to.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject(error);
      }
  })
}

function getCcRecipients(item) {
    return new Office.Promise(function (resolve, reject) {
      try {
        item.cc.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject(error);
      }
  })
}

function getBccRecipients(item) {
  return new Office.Promise(function (resolve, reject) {
    try {
      item.bcc.getAsync(function (asyncResult) {
            resolve(asyncResult.value);
        });
    }
    catch (error) {
        reject(error);
    }
})
}

function btnSendClick() {
  Office.context.ui.messageParent(true);
}

function btnCancelClick() {
  Office.context.ui.messageParent(false);
}

function processMessage(event) {
  let allow = event.message ? true : false;
  let item = Office.context.mailbox.item;

  if (!allow)
  {
    item.close();
    sendEvent.completed({ allowEvent: false });
  } 
  else 
  {
    item.subject.getAsync(
      {},
      function(result) {
          let subject = result.value;
          item.subject.setAsync('[Secure] ' + subject, function() {
            sendEvent.completed({ allowEvent: true });
          })
      })
  }
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
g.validateRecipients = validateRecipients;
g.btnSendClick = btnSendClick;
g.btnCancelClick = btnCancelClick;
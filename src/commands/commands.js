/*
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */

var dialog;
var sendEvent;

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

function validateRecipients(event) {
  sendEvent = event;

  let item = Office.context.mailbox.item;
  let toRecipientPromise = getToRecipients(item);
  let ccRecipientPromise = getCcRecipients(item);
  let bccRecipientPromise = getBccRecipients(item);

  Promise.all([toRecipientPromise, ccRecipientPromise, bccRecipientPromise]).then((promises) => {
    let hasExternal = false;

    const combinedRecipients = [...promises[0],...promises[1],...promises[2]];

    for (let i = 0; i < combinedRecipients.length; i++) {
      if (combinedRecipients[i].recipientType === "externalUser") {
        hasExternal = true;
        break;
      }
    }

    if (hasExternal) {
      Office.context.ui.displayDialogAsync('https://localhost:3000/validate.html', { height: 18, width: 30, promptBeforeOpen: false}, //TODO relative URLs are apparently not supported. Add non-hard-coded server path 
      function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        
        const topRecipients = combinedRecipients.slice(0,3).map((item) => {return item.emailAddress;}); //TODO configurable slice length
        const recipientString = JSON.stringify(topRecipients);

        dialog.messageChild(recipientString);
      });
    } 
    else {
      event.completed({ allowEvent: true });
    }
  });
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
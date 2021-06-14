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

  var item = Office.context.mailbox.item;
  item.to.getAsync(function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({ allowEvent: true });
      }
      else {
          var recipients = asyncResult.value;
          var hasExternal = false;

          for (var i = 0; i < recipients.length; i++) {
            if (recipients[i].recipientType === "externalUser") {
              hasExternal = true;
              break;
            }
          }

          if (hasExternal) {
            Office.context.ui.displayDialogAsync('https://localhost:3000/validate.html', { height: 12, width: 20, promptBeforeOpen: false},
            function (result) {
              dialog = result.value;
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            });
          } 
          else {
            event.completed({ allowEvent: true });
          }
      }
  });
}

function btnSendClick() {
  Office.context.ui.messageParent(true);
}

function btnCancelClick() {
  Office.context.ui.messageParent(false);
}

function processMessage(event) {
  var allow = event.message ? true : false;

  if (!allow)
  {
    var item = Office.context.mailbox.item;
    item.close();
  }

  sendEvent.completed({ allowEvent: allow });
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
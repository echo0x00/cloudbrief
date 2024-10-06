/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Office.onReady(() => {
//   // Office.context.mailbox.item.notificationMessages.replaceAsync("action", "read");
//   // displayNotification("Это тестовое уведомление.");
// });

Office.initialize = function (reason) {
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
};

function displayNotification(message) {
  Office.context.mailbox.item.notificationMessages.addAsync(
    "info",
    {
      type: "informationalMessage",
      message: message,
      icon: "icon-16",
      persistent: true,
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Ошибка добавления уведомления: " + asyncResult.error.message);
      }
    }
  );
}

function action(event) {
  // displayNotification("Просто тестовое уведомление.");

  const item = Office.context.mailbox.item;

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    item.body.getAsync("html", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        body = result.value;

        if (body.indexOf("фишинг") > -1) {
          var url = new URI("dialog.html").absoluteTo(window.location).toString();
          console.log(url);
          Office.context.ui.displayDialogAsync(url, { height: 30, width: 50, displayInIframe: true }, () => {
            // event.completed();
          });
        }
      }
    });
  }
}

function itemChanged(eventArgs) {
  const item = Office.context.mailbox.item;

  console.log(item.itemType);

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    item.body.getAsync("html", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        body = result.value;

        if (body.indexOf("фишинг") > -1) {
          var url = new URI("dialog.html").absoluteTo(window.location).toString();
          console.log(url);
          Office.context.ui.displayDialogAsync(url, { height: 30, width: 50, displayInIframe: true }, () => {
            event.completed();
          });
        }
      }
    });
  }
}

Office.actions.associate("action", action);

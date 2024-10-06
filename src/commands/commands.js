/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
  // console.log("Item Change event registered.");
});

Office.initialize = function (reason) {};

function displayError(message) {
  Office.context.mailbox.item.notificationMessages.removeAsync("err");
  Office.context.mailbox.item.notificationMessages.addAsync(
    "err",
    {
      type: "errorMessage",
      message: message,
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Ошибка добавления уведомления: " + asyncResult.error.message);
      }
    }
  );
}

function displayNotification(message) {
  Office.context.mailbox.item.notificationMessages.removeAsync("info");
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
  const item = Office.context.mailbox.item;

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    item.body.getAsync("html", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        body = result.value;

        if (body.indexOf("фишинг") > -1 || item.from.emailAddress.indexOf("cloud.ru") == -1) {
          setTimeout(() => {
            displayError("Письмо похоже на фишинг");
            event.completed();
          }, 5000);

          Office.context.ui.displayDialogAsync(
            "https://echo0x00.github.io/cloudbrief/src/dialog/dialog.html?from=" + item.from.emailAddress,
            { height: 40, width: 50, displayInIframe: true },
            function (asyncResult) {}
          );
        } else {
          displayNotification("Фишинг не обнаружен.");
          event.completed();
        }
      }
    });
  }
}

// function itemChanged(eventArgs) {
//   console.log("Another email message selected");

//   const item = Office.context.mailbox.item;

//   console.log(item.itemType);

//   if (item.itemType === Office.MailboxEnums.ItemType.Message) {
//     item.body.getAsync("html", function (result) {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         body = result.value;

//         if (body.indexOf("фишинг") > -1 || body.indexOf("asd") > -1) {
//           Office.context.ui.displayDialogAsync(
//             "https://echo0x00.github.io/cloudbrief/src/dialog/dialog.html",
//             { height: 30, width: 50, displayInIframe: true },
//             () => {
//                event.completed();
//             }
//           );
//         }
//       }
//     });
//   }
// }

Office.actions.associate("action", action);

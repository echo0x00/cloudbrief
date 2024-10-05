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
  Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
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
    item.body.setAsync("<p>New HTML content for the email body.</p>", { coercionType: Office.CoercionType.Html });
  }

  event.completed();
}

function itemChanged(eventArgs) {
  const item = Office.context.mailbox.item;
  let subject = "test";

  displayNotification(item);

  // item.subject.getAsync({ asyncContext: item }, function (asyncResult) {
  //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //     subject = asyncResult.value;
  //     console.log(`Открыто письмо с темой: ${subject}`);
  //     // Добавьте здесь код для обработки вашего события
  //   } else {
  //     console.error("Ошибка получения темы", asyncResult.error);
  //   }
  // });

  // if (item.itemType === Office.MailboxEnums.ItemType.Message) {
  //   item.body.setAsync(
  //     "<p>New HTML content for the email body.</p>",
  //     { coercionType: Office.CoercionType.Html },
  //     function (asyncResult) {
  //       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
  //         console.error("Failed to set body:", asyncResult.error.message);
  //       } else {
  //         console.log("Body content set successfully.");
  //       }
  //     }
  //   );
  // }
}

// function getGlobal() {
//   return typeof self !== "undefined"
//     ? self
//     : typeof window !== "undefined"
//       ? window
//       : typeof global !== "undefined"
//         ? global
//         : undefined;
// }

// var g = getGlobal();

// // the add-in command functions need to be available in global scope
// g.action = action;

// Register the function with Office.
Office.actions.associate("action", action);

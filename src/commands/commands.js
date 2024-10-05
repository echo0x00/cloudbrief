/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {});

Office.initialize = function (reason) {
  Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
};

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // const message = {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: "привет",
  //   icon: "Icon.80x80",
  //   persistent: true,
  // };

  // // Show a notification message.
  // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // // Be sure to indicate when the add-in command function is complete.

  const item = Office.context.mailbox.item;

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    var dialogOptions = { width: 40, height: 60, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
      loginDialog = result.value;
      loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
    });

    item.body.setAsync(
      "<p>New HTML content for the email body.</p>",
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set body:", asyncResult.error.message);
        } else {
          console.log("Body content set successfully.");
        }
      }
    );
  }
  // event.completed();
}

function dialogClosed(arg) {
  loginDialog = null;
  configEvent.completed();
  configEvent = null;
}

function itemChanged(eventArgs) {
  const item = Office.context.mailbox.item;
  let subject = "test";

  console.log(`Открыто письмо с темой: ${subject}`);

  item.subject.getAsync({ asyncContext: item }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      subject = asyncResult.value;
      console.log(`Открыто письмо с темой: ${subject}`);
      // Добавьте здесь код для обработки вашего события
    } else {
      console.error("Ошибка получения темы", asyncResult.error);
    }
  });

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    item.body.setAsync(
      "<p>New HTML content for the email body.</p>",
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set body:", asyncResult.error.message);
        } else {
          console.log("Body content set successfully.");
        }
      }
    );
  }
}

// Register the function with Office.
Office.actions.associate("action", action);

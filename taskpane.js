/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */


Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Das Skript wurde erfolgreich geladen!");

    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      console.log("Event handler added.");
    });

    run();
    
  }
});


async function run() {
  // Get a reference to the current message
  // *const item = Office.context.mailbox.item;
  // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Betreff:</b> <br/>" + item.subject;
  // document.getElementById("item-attachments").innerHTML = "<b>Anzahl Anlagen:</b> <br/>" + item.attachments.length;
  // document.getElementById("item-from").innerHTML = "<b>Von:</b> <br/>" + item.from.emailAddress;
  // document.getElementById("item-to").innerHTML = "<b>An:</b> <br/>" + item.to[0]["emailAddress"];
  // document.getElementById("item-guid").innerHTML = "<b>GUID:</b> <br/>" + item.itemId;
  //const subject = item.subject;
  //const numberAttachments = item.attachments.length;
  //const from = item.from.emailAddress;
  //const toRecipient = item.to[0]["emailAddress"];

  // *const mailGUID = encodeURIComponent(item.itemId);
  // *const subject = encodeURIComponent(item.subject);
  // *const date = encodeURIComponent(item.dateTimeCreated.toISOString());

  // *console.log("Halli hallooo");
  // *console.log("Datum = " + date);

  Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;      
    }

    asyncResult.value.forEach((item, index) => {
        let mailGuidArray = [];
        let mailGuidKey = "mailGUID" + (parseInt(++index));
        let mailGuidValue = encodeURIComponent(item.itemId);
        mailGuidArray.push({[mailGuidKey]: mailGuidValue});
        let joinedArray = "";

        for (const arrayItem of mailGuidArray) {
          let keys = Object.keys(arrayItem);
          for (const key of keys) {
            joinedArray += "&" + key + "=" + mailGuidArray[key];
          }
        }

        const httpString = document.getElementById("joinedArrayDiv");
        httpString.innerText = joinedArray;
    });
});

  const httpString = document.getElementById("joinedArrayDiv");
  httpString.innerText = joinedArray;

  document.getElementById("powerapps-iframe").setAttribute("src", "https://apps.powerapps.com/play/9c9d578c-eb09-4c0c-81a4-9df8ecca7650?tenantId=b9d0cec0-1f3e-4a91-8ec5-7341d8dbe1e0&source=iframe&mailGuid=" + mailGUID + "&subject=" + subject + "&mailDate=" + date);
}

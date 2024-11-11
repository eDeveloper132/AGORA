/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
let currentComposeId = null;
// Add start-up logic code here, if any.
Office.onReady();

function onNewMessageComposeHandler(event) {
  onNewComposeWindowLaunch();
}

function onNewComposeWindowLaunch() {
  emailSent = false;
  currentComposeId = Office.context.mailbox.item.conversationId;
  storageKey = `attachments-${currentComposeId}`;
  setInLocalStorage(storageKey, JSON.stringify([]));
}

function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}

function cleanupAttachments(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.removeItem(myPartitionKey + key);
  } else {
    return localStorage.removeItem(key);
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);

/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
let storageKey = null;
let emailSent = false;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox?.item;
};

function onNewComposeWindow() {
  emailSent = false;
  currentComposeId = Office.context.mailbox.item.itemId;
  storageKey = "attachments";

  if (!getFromLocalStorage("attachments")) {
    setInLocalStorage("attachments", JSON.stringify([]));
  }
}

async function uploadFileToRoom(guestRID, event, roomId, userId, file, accessToken) {
  const url = globalConfig.backendBaseUrl + `/agora-api/api/v1/rooms/${roomId}/files`;

  const formData = new FormData();

  const metadata = {
    type: 5,
    subtype: 2,
    permissions: {
      items: [
        {
          rid: userId,
          type: 2,
          subtype: 2,
          permission: 5,
          level: 5,
          inheritance: 0,
        },
      ],
    },
    name: file.name,
    expireDate: file.expire,
  };
  formData.append("Resource", JSON.stringify(metadata));
  try {
    const byteCharacters = atob(file.data);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: file.type });
    formData.append("file", blob, file.name);
  } catch (error) {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      error.toString(),
      { coercionType: Office.CoercionType.Text },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
        }
      }
    );
  }

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json, text/plain, */*",
      },
      body: formData,
    });

    if (response.ok) {
      const data = await response.json();
      //return data;
      const urlGuest = globalConfig.backendBaseUrl + `/agora-api/api/v1/files/${data.rid}/guest`;
      const formDataGuest = new FormData();
      var raw;
      if (guestRID == 0) {
        raw = JSON.stringify({
          expireDate: `${file.expire}`,
          type: 2,
          subtype: 3,
          authenticationMethod: 0,
          mobile: "",
          language: "en_US",
          permission: `${file.permission}`,
          country: "ch",
          email: `${data.rid}@${data.rid}.com`,
          guestAccessPropagate: true,
          loginInfo: true,
          token: true,
        });
      } else {
        raw = JSON.stringify({
          rid: `${guestRID}`,
          expireDate: `${file.expire}`,
          type: 2,
          subtype: 3,
          authenticationMethod: 0,
          mobile: "",
          language: "en_US",
          permission: `${file.permission}`,
          country: "ch",
          guestAccessPropagate: true,
        });
      }

      const responseGuest = await fetch(urlGuest, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json, text/plain, */*",
          "Content-Type": "application/json",
        },
        body: raw,
        redirect: "follow",
      });

      if (responseGuest.ok) {
        const responseGuestJSON = await responseGuest.json();
        await new Promise((r) => setTimeout(r, 3000));
        responseGuestJSON.fid = data.rid;
        return responseGuestJSON;
      } else {
        const errorData = await responseGuest.text();
        throw new Error(`File upload failed: ${errorData}`);
      }
    } else {
      const errorData = await response.text();
      Office.context.mailbox.item.body.setSelectedDataAsync(
        errorData,
        { coercionType: Office.CoercionType.Text },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
          }
        }
      );
      console.error("File upload error response:", errorData);
      // event.completed({ allowEvent: false });
      throw new Error(`File upload failed: ${errorData}`);
    }
  } catch (error) {
    console.error("An error occurred during file upload:", error);
    await addBodyOnSend(event, error.message);
    throw error;
  }
}
// Entry point for AGORA add-in before send is allowed.
function Upload(event) {
  mailboxItem.body.getAsync("text", { asyncContext: event }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
    }
    UploadOnlyOnSendCallBack(asyncResult);
  });
}

function addBodyOnSend(event, variable) {
  return new Promise((resolve, reject) => {
    mailboxItem.body.setSelectedDataAsync(variable, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}
function addLinkOnSend(event, variable, displayText) {
  var hyperlink = '<br><a href="' + variable + '">' + displayText + "</a><br>";
  return new Promise((resolve, reject) => {
    mailboxItem.body.setSelectedDataAsync(hyperlink, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}

async function UploadOnlyOnSendCallBack(event) {
  var globalAttachments = [];
  globalAttachments = await getAttachments();
  const message = getFromLocalStorage("myKey1");
  currentComposeId = Office.context.mailbox.item.conversationId;
  storageKey = `attachments-${currentComposeId}`;
  let attachments = JSON.parse(getFromLocalStorage(storageKey) || "[]");
  console.log(`${attachments.length}`);
  console.log(`${globalAttachments.length}`);
  //  let attachments_internal = Office.context.mailbox.item.attachments;
  if (attachments.length == 0) {
    try {
      console.log("Email body updated successfully.");
    } catch (error) {
      console.error("Failed to update email body: " + error.message);
      // Handle the error here
    }
    event.asyncContext.completed({ allowEvent: true });
  }
  if (attachments.length > 0) {
    var final = event.value.toString();
    var guestRID = 0;
    var firstlink = "";
    var link = "";
    var found = false;

    for (let attachment of attachments) {
      console.log(attachment.name);
      if (globalAttachments.length > 0) {
        for (let i = 0; i < globalAttachments.length; i++) {
          console.log(`locals:${attachment.name}`);
          console.log(`global:${globalAttachments[i].name}`);
          const attachment_internal = globalAttachments[i];
          if (attachment.name == attachment_internal.name) {
            found = true;
            await Office.context.mailbox.item.removeAttachmentAsync(attachment_internal.id);
            break;
          }
        }
      }
      if (found == false) {
        continue;
      }
      found = false;
      const roomId = getFromLocalStorage("access-roomId");
      const access = getFromLocalStorage("access");
      const userId = getFromLocalStorage("access-userId");
      try {
        console.log("Email body updated successfully.");
      } catch (error) {
        console.error("Failed to update email body: " + error.message);
      }

      try {
        const regex = /file=[0-9]+/i;
        const result = await uploadFileToRoom(guestRID, event.asyncContext, roomId, userId, attachment, access);
        if (guestRID == 0) {
          firstlink = result.password;
          link = firstlink;
          guestRID = result.rid;
          console.log("first link  " + link);
        } else {
          decodedlink = decodeURIComponent(firstlink);
          console.log("encoded link  " + firstlink);
          console.log("decoded link  " + decodedlink);
          link = decodedlink.replace(regex, `file=${result.fid}`);
          console.log("second link  " + link);
        }
        await addLinkOnSend(event.asyncContext, link, attachment.name);
      } catch (error) {
        console.error("Failed to update email body: " + error.message);
        await addBodyOnSend(event.asyncContext, error.message);
        event.asyncContext.completed({ allowEvent: false });
      }
    }
  }
  //clean the attachments

  cleanupAttachments(storageKey);
  // Allow send.
  event.asyncContext.completed({ allowEvent: true });
}

async function openDialog(event) {
  const ready = await environmentReady();

  if (!ready) {
    Office.context.ui.displayDialogAsync(
      globalConfig.serverBaseUrl + "/taskpane/taskpane.html",
      { width: 50, height: 50, displayInIframe: true },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Error: ${asyncResult.error.message}`);
          Office.context.mailbox.item.notificationMessages.addAsync(
            "openTaskPaneWarning",
            {
              type: "informationalMessage",
              message: `Error: ${asyncResult.error?.message}`,
              icon: "icon16",
              persistent: false,
            },
            function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Warning message displayed successfully.");
              } else {
                console.error(`Error displaying warning message: ${asyncResult.error.message}`);
              }
            }
          );
        }
        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          console.log("Dialog closed.");
          event.completed();
        });
      }
    );
  } else {
    Office.context.ui.displayDialogAsync(
      globalConfig.serverBaseUrl + "/commands/dialog.html",
      { width: 50, height: 50, displayInIframe: true },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Error: ${asyncResult.error.message}`);
        }
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (message) {
          addAttachment(event, JSON.parse(message.message));
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          console.log("Dialog closed.");
          event.completed();
        });
      }
    );
  }
}

function addAttachment(event, file) {
  var attachment = file;
  currentComposeId = Office.context.mailbox.item.conversationId;
  storageKey = `attachments-${currentComposeId}`;
  var aa = JSON.parse(getFromLocalStorage(storageKey));
  if (!aa) {
    aa = [];
  }
  aa.push(attachment);

  try {
    setInLocalStorage(storageKey, JSON.stringify(aa));
  } catch (error) {
    addBodyOnSend(event, error.message);
    localStorage.clear();
  }
  const attachmentFile = {
    type: Office.MailboxEnums.AttachmentType.File,
    name: attachment.name,
    content: "This is an example attachment content.",
  };

  Office.context.mailbox.item.addFileAttachmentFromBase64Async(file.data, attachmentFile.name);

  console.log("Attachments stored in sessionStorage:");
  return;
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
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}

function cleanupAttachments(key) {
  const myPartitionKey = Office.context.partitionKey;
  if (myPartitionKey) {
    return localStorage.removeItem(myPartitionKey + key);
  } else {
    return localStorage.removeItem(key);
  }
}

async function environmentReady() {
  const accessToken = getFromLocalStorage("access");
  if (accessToken == null) {
    return false;
  }
  const roomId = getFromLocalStorage("access-roomId");
  if (roomId == null) {
    Office.context.mailbox.item.notificationMessages.addAsync(
      "openTaskPaneWarning",
      {
        type: "informationalMessage",
        message:
          "Your Personal Room is not created or deleted\nPlease check and create a Personal Room if needed then login again.",
        icon: "icon16",
        persistent: false,
      },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Warning message displayed successfully.");
        } else {
          console.error(`Error displaying warning message: ${asyncResult.error.message}`);
        }
      }
    );
    return false;
  }
  const finalResult = await checkRoom(accessToken);
  return finalResult;
}

async function checkRoom(accessToken) {
  try {
    const response = await fetch(`${globalConfig.backendBaseUrl}/agora-api/api/v1/user/settings/`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json, text/plain, */*",
      },
    });

    const data = await response.json();
    if (response.ok && data.roomId && data.roomPermission >= 4) {
      const roomId = data.roomId;
      setInLocalStorage("access-roomId", roomId);
      return true;
    } else {
      Office.context.mailbox.item.notificationMessages.addAsync(
        "openTaskPaneWarning",
        {
          type: "informationalMessage",
          message:
            "Your Personal Room is not created or deleted\nPlease check and create a Personal Room if needed then login again.",
          icon: "icon16",
          persistent: false,
        },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Warning message displayed successfully.");
          } else {
            console.error(`Error displaying warning message: ${asyncResult.error.message}`);
          }
        }
      );
      return false;
    }
  } catch (error) {
    mailboxItem.body.setSelectedDataAsync(error.message, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
      }
    });
    return false;
  }
}
function getAttachments() {
  return new Promise(function (resolve, reject) {
    var Attachments = [];
    Office.context.mailbox.item.getAttachmentsAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Update the variable with the attachments
        Attachments = asyncResult.value;
        resolve(Attachments);
      } else {
        reject("Error getting attachments: " + asyncResult.error.message);
      }
    });
  });
}
Office.actions.associate("openDialog", openDialog);

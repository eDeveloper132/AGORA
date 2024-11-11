/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */

/* global Office */

Office.onReady((info) => {
  /*
  if (info.host === Office.HostType.Outlook) {
    document.addEventListener("DOMContentLoaded", function () {
      document.getElementById("auth-upload-form").addEventListener("submit", async function (event) {
        event.preventDefault();
        Office.context.ui.message("File selected:");
        const username = document.getElementById("username").value;
        const password = document.getElementById("password").value;
        const trustRoomName = document.getElementById("trustRoomName").value;
        const roomId = document.getElementById("roomId").value;
        const fileInput = document.getElementById("fileInput");
        const file = fileInput.files[0];

        const output = document.getElementById("output");
        output.textContent = "Authenticating...";

        try {
          const accessToken = await authenticate(username, password, trustRoomName);
          output.textContent = "Authenticated successfully. Uploading file...";

          const result = await uploadFileToRoom(roomId, file, accessToken);
          output.textContent = `File uploaded successfully:\n${JSON.stringify(result, null, 2)}`;
        } catch (error) {
          output.textContent = `Operation failed: ${error.message}`;
        }
      });
    });
  }
  */
});

async function authenticate(username, password, trustRoomName) {
  const url = globalConfig.backendBaseUrl + "/agora-api/api/v1/auth";
  const clientId = "outlook";
  const clientSecret = "secret";
  const grantType = "password";

  const formData = new URLSearchParams();
  formData.append("grant_type", grantType);
  formData.append("client_id", clientId);
  formData.append("client_secret", clientSecret);
  formData.append("username", username);
  formData.append("password", password);
  formData.append("community", trustRoomName);

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Accept: "application/json, text/plain, */*",
      },
      body: formData.toString(),
    });

    if (response.ok) {
      const data = await response.json();
      return data;
    } else {
      const errorData = await response.text();
      console.error("Authentication error response:", errorData);
      throw new Error(`Authentication failed: ${errorData}`);
    }
  } catch (error) {
    console.error("An error occurred during authentication:", error);
    throw error;
  }
}

async function performAuthentication(event) {
//  document.getElementById("auth-upload-form").addEventListener("submit", async function (event) {
//    event.preventDefault();
    const username = document.getElementById("username").value;
    const password = document.getElementById("password").value;
    const trustRoomName = document.getElementById("trustRoomName").value;
    const output = document.getElementById("output");
    output.textContent = "Authenticating...";

    try {
      const data = await authenticate(username, password, trustRoomName);
      const accessToken = data.access_token;
      let roomId = null;
      if (data.settings) {
        if (data.settings.roomId && data.settings.roomPermission >= 4) {
          roomId = data.settings.roomId;
          console.debug('Room ID', roomId);
          setInLocalStorage("access-roomId", roomId);
        }  
        setInLocalStorage("access-userId", data.settings.userId);
      }
      output.textContent = "Authenticated successfully. \nThe credentials will be saved for next time..";
      setInLocalStorage("access", accessToken);
      if (!roomId) {
        output.textContent = "No personal room found. \nPlease create one then login again";
      }
      //return accessToken;
    } catch (error) {
      output.textContent = `Operation failed: ${error.message}`;
    }
//  });
  //event.completed();
}

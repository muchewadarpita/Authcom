/* eslint-disable office-addins/no-office-initialize */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// import "isomorphic-fetch";
// import { DeviceCodeCredential, ClientSecretCredential } from "@azure/identity";
// import { Client } from "@microsoft/microsoft-graph-client";
// import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
const { PublicClientApplication } = require("@azure/msal-browser");

const clientId = "b3884fb9-2389-4290-a3f9-02dfd4813955";
const authority = "https://login.microsoftonline.com/6d084a41-2a0b-4141-a5fa-1eb080f65327";
const redirectUri = "http://localhost:3000/callback";
var accessToken = null;
const settings = {
  clientId: "b3884fb9-2389-4290-a3f9-02dfd4813955",
  clientSecret: "YeR8Q~uVqavFXOUHdJ-4sqorkMcmgcQGHU_aWbOB",
  tenantId: "6d084a41-2a0b-4141-a5fa-1eb080f65327",
  authTenant: "common",
  graphUserScopes: ["user.read", "mail.read", "openid", "offline_access"],
};

// import { PublicClientApplication, InteractionType } from "@azure/msal-browser";

Office.onReady(async () => {
  // login();
  console.log("ronReadyes---------->>>");
  // login();
  console.log("ronReadyes---------->>>");
  Office.context.mailbox.getCallbackTokenAsync(
    {
      isRest: true,
    },
    async function (asyncResult) {
      console.log("======onReadyonReadyonReadyonReadyonReady : ");
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        authToken = asyncResult.value;
        console.log("authToken----->>> ");
        console.log(authToken);
        if (authToken != "" && authToken != undefined) {
          var roamingSettings = Office.context.roamingSettings;
          // Store the token
          var token = authToken; // Replace with your actual token
          roamingSettings.set("accessToken", token);
          // Save changes
          roamingSettings.saveAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error("Failed to store token:", result.error.message);
            }
          });
        }
      } else {
        console.error("Failed to retrieve message ID: " + asyncResult.error.message);
      }
    }
  );
  // Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, onSendHandler);
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, onSendHandler);
});

async function login() {
  const msalConfig = {
    auth: {
      clientId: "b3884fb9-2389-4290-a3f9-02dfd4813955",
      redirectUri: "https://localhost:3000",
      authority: "https://login.microsoftonline.com/6d084a41-2a0b-4141-a5fa-1eb080f65327",
      grantType: "authorization_code",
      scope: "https://graph.microsoft.com/User.Read",
    },
  };

  const msalInstance = await new PublicClientApplication(msalConfig);
  await msalInstance.handleRedirectPromise();
  try {
    const response = await msalInstance.loginPopup();
    console.log("Information : responseresponseresponseresponseresponse-->");
    console.log(response.accessToken);
    // if (response.accessToken != null && response.accessToken != "") {
    accessToken = response.accessToken;
    // }
    // Handle login success
  } catch (error) {
    // Handle login failure
    console.error("Login error:", error);
  }
}

window.getUserTokenAsync = function getUserTokenAsync() {
  return accessToken;
};

function onItemSend(eventArgs) {
  // Perform your desired action here, for example, displaying a notification or executing a function.
  console.log("Email sent successfully!");

  // You can also access the sent email item using eventArgs.getCallbackTokenAsync() and make further changes if needed.
}

// Function to add the event handler when the add-in starts
function addEventHandlers() {
  // Add the ItemSend event handler
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, onItemSend, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("ItemSend event handler added successfully.");
    } else {
      console.error("Error adding ItemSend event handler:", result.error.message);
    }
  });
}

function onSendHandler() {
  // Add the ItemSend event handler
  console.log("00000----->>");
  Office.context.mailbox.item.sendAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("1111111----->>");
      // The email was sent successfully
      // Include a unique identifier or token in the email body to identify the response later
      // For example, you can add a custom header or a specific text in the email body
    }
  });
}

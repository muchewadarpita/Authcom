/* eslint-disable office-addins/no-office-initialize */
/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */

var item;
var fromMail = null;
var userName = null;
var toReceiptList = [];
var bccReceiptList = [];
var ccReceiptList = [];
var subject = null;
var messageId = null;
var itemId = null;
var randomstring = null;
var body = null;
var PromiseList = [];
var attachmentID = null;
var acutalBody = null;
var authToken = "";
var getWayName = "OUTLOOK";
var emailUrl = `https://outlook.office.com/api/v2.0/me/messages`;
var webHookUrl = "https://gessa.io/authcomm-bff/";
// var webHookUrl = "http://localhost:40008/";

function callback(result) {}

Office.initialize = function (reason) {
  console.log("Received email detected000.");
  item = Office.context.mailbox.item;
  console.log(item);
  const options = {
    asyncContext: {
      currentItem: item,
    },
  };
  item.getAttachmentsAsync(options, callback);
  if (reason === Office.InitializationReason.DocumentReady) {
    // Initialize the interval (in milliseconds)
    var interval = 60000; // 1 minute
    console.log("Received email detecte11111.");
    // Set up a function to be executed periodically
    var intervalId = setInterval(checkForNewEmails, interval);
  }
};
Office.context.mailbox.item.addHandlerAsync(
  Office.EventType.ItemChanged,
  {
    itemType: Office.MailboxEnums.ItemType.Message,
  },
  function (eventArgs) {
    console.log("Received email detecte22222111111.");
    if (eventArgs.isNew) {
      // Handle the event for new emails received
      // This code will run when a new email arrives in the inbox
    }
  }
);

function checkForNewEmails() {
  console.log("Received email detecte22222.");
  var inbox = Office.context.mailbox.folders.getFolderType(Office.MailboxEnums.FolderType.Inbox);

  inbox.getCallbackTokenAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Received email detecte3333333Z.");
      var callbackToken = result.value;

      // Use callbackToken to make Microsoft Graph API requests
      // Example: Call Microsoft Graph API to retrieve email count or other data
      // Make sure to handle authentication and permissions properly
      // ...
    } else {
      console.error("Error getting callback token:", result.error);
    }
  });
}

function generateRandomUniquNo(prefix, length) {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXTZ0123456789";
  string_length = prefix ? length - prefix.length : length;
  var token = "";
  for (var i = 0; i < string_length; i++) {
    token += chars[Math.floor(Math.random() * chars.length)];
  }
  return token;
}

function jsonEscape(str) {
  return str.replace(/\n/g, "\\\\n").replace(/\r/g, "\\\\r").replace(/\t/g, "\\\\t");
}

async function getFileInfo(authToken1) {
  var roamingSettings = Office.context.roamingSettings;
  // Retrieve the token
  var authToken2 = roamingSettings.get("accessToken");
  var request = new XMLHttpRequest();
  let url = await emailUrl;
  console.log("------>>>>");
  console.log(authToken2);
  console.log("Function1");
  console.log(url);
  console.log("getFileInfo");
  request.open("GET", url, false);
  request.setRequestHeader("Content-Type", "application/json");
  request.setRequestHeader("Authorization", "Bearer " + authToken2);
  request.onreadystatechange = function () {
    console.log("2222111IN FileInfo----->");
    if (request.readyState === 4) {
      if (request.status === 200) {
        var response = JSON.parse(request.responseText);
        console.log("FileInforesponse----->", response.value[0].Id);
        console.log("FileInforesponse----->", response.value[0].ConversationId);
        getEmailSMTPInfo(response.value[0].Id, authToken1, response.value[0].ConversationId);
        // getEmailSMTPInfo(messageId, authToken1, emailThread, conversationId);
        // getEmailThread(response.value[0].Id, authToken1, response.value[0].ConversationId);
        // getAttachments(response.value[0].Id, authToken1, response.value[0].ConversationId);
      } else {
        console.log("FileInfo1error----->", request.status);
      }
    } else {
      console.log("FileInfo2error----->", request.responseText);
    }
  };
  await request.send();
}

Office.onReady(async () => {
  Office.context.mailbox.getCallbackTokenAsync(
    {
      isRest: true,
    },
    async function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        authToken = asyncResult.value;
        console.log("authToken----->>> ");
        console.log(authToken);
        if (authToken != "" && authToken != undefined) {
          var roamingSettings = Office.context.roamingSettings;
          // Store the token
          var token = authToken; // Replace with your actual token
          roamingSettings.set("accessToken", token);
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
});

async function getAttachments(messageId, authToken1, smtpDetails, conversationId) {
  var roamingSettings = Office.context.roamingSettings;
  // Retrieve the token
  var authToken2 = roamingSettings.get("accessToken");
  var request = new XMLHttpRequest();
  console.log("V222222");
  var url = emailUrl + "/" + messageId + "/attachments/";
  request.open("GET", url, false);
  console.log("Function4");
  console.log(url);
  console.log("getAttachments");
  request.setRequestHeader("Content-Type", "application/json");
  request.setRequestHeader("Authorization", "Bearer " + authToken2);
  request.onreadystatechange = async function () {
    console.log("2222111IN FileInfo----->");
    if (request.readyState === 4) {
      // console.log("4444444 FileInfo----->");
      // console.log("responseText----->", request.responseText);
      if (request.status === 200) {
        var response = JSON.parse(request.responseText);
        var attachmentList = [];
        // Process the response data here
        for (let i = 0; i < response.value.length; i++) {
          // byteArray[i] = binaryString.charCodeAt(i);
          if (
            response.value[i].Name != "Outlook-xulcglg2" &&
            response.value[i].ContentType != "image/svg+xml"
            // &&
            // response.value[i].ContentBytes !=
            //   "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"
          ) {
            console.log("1FileInforesponse----->", response.value[i].ContentBytes);
            console.log("2FileInforesponse----->", response.value[i].Name);
            console.log("3FileInforesponse----->", response.value[i].ContentType);
            attachmentList.push({
              file_blob: response.value[i].ContentBytes,
              file_name: response.value[i].Name,
              mime_Type: response.value[i].ContentType,
              file_size: response.value[i].Size,
            });
          }
        }
        console.log("afterEmailSend---->>>>");
        console.log("afterEmailSend---->>>>");
        afterEmailSend({
          receiver: toReceiptList.length > 0 ? toReceiptList[0].emailAddress : "",
          bcc: bccReceiptList > 0 ? bccReceiptList[0] : "",
          cc: ccReceiptList > 0 ? ccReceiptList[0] : "",
          user_name: userName,
          sender: fromMail,
          subject: subject,
          msg_id: messageId,
          body: jsonEscape(acutalBody),
          cert_id: randomstring,
          sending_date_time: new Date().toLocaleString(),
          delivered_date_time: new Date().toLocaleString(),
          gateway_name: "OUTLOOK",
          attachments: attachmentList,
          // email_thread: emailThread,
          email_thread: [],
          smtp_details: smtpDetails,
          conversation_id: conversationId,
        });
        // byteArrayToFile(attachmentList[0].ContentBytess, attachmentList[0].Name, attachmentList[0].ContentType);
        // await getEmailThread(messageId, authToken1, smtpDetails, conversationId, attachmentList);
      } else {
        // Handle the error here
        console.log("FileInfo1error----->", request.status);
      }
    } else {
      console.log("FileInfo2error----->", request.responseText);
    }
  };

  console.log("00json Request ----->");
  // var payloadJson = JSON.stringify(json);
  await request.send();
}

async function manageEmailThread(messageId, authToken1, smtpDetails, conversationId, attachmentList, emailThread) {
  var roamingSettings = Office.context.roamingSettings;
  // Retrieve the token
  var authToken2 = roamingSettings.get("accessToken");
  for (let i = 0; i < emailThread.length; i++) {
    var request = new XMLHttpRequest();
    console.log("V222222");
    var url = emailUrl + "/" + emailThread[i].Id + "/attachments/";
    request.open("GET", url, false);
    console.log("Function44");
    console.log(url);
    console.log("getAttachments44");
    request.setRequestHeader("Content-Type", "application/json");
    request.setRequestHeader("Authorization", "Bearer " + authToken2);
    request.onreadystatechange = function () {
      console.log("2222111IN FileInfo----->");
      if (request.readyState === 4) {
        // console.log("4444444 FileInfo----->");
        // console.log("responseText----->", request.responseText);
        if (request.status === 200) {
          var response = JSON.parse(request.responseText);
          var attachmentList = [];
          // Process the response data here
          for (let i = 0; i < response.value.length; i++) {
            // byteArray[i] = binaryString.charCodeAt(i);
            if (
              response.value[i].Name != "Outlook-xulcglg2" &&
              response.value[i].ContentType != "image/svg+xml" &&
              response.value[i].ContentBytes !=
                "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"
            ) {
              console.log("1FileInforesponse----->", response.value[i].ContentBytes);
              console.log("2FileInforesponse----->", response.value[i].Name);
              console.log("3FileInforesponse----->", response.value[i].ContentType);
              attachmentList.push({
                file_blob: response.value[i].ContentBytes,
                file_name: response.value[i].Name,
                mime_Type: response.value[i].ContentType,
                file_size: response.value[i].Size,
              });
            }
          }
          console.log("manageEmailThreadafterEmailSend---->>>>");
          console.log(attachmentList);
          emailThread[i]["attachments"] = attachmentList;
          // console.log(json);
          // return emailThread;
        } else {
          // Handle the error here
          console.log("FileInfo1error----->", request.status);
        }
      } else {
        console.log("FileInfo2error----->", request.responseText);
      }
    };

    console.log("111json Request ----->");
    // var payloadJson = JSON.stringify(json);
    await request.send();
  }
  // Final
  afterEmailSend({
    receiver: toReceiptList.length > 0 ? toReceiptList[0].emailAddress : "",
    bcc: bccReceiptList > 0 ? bccReceiptList[0] : "",
    cc: ccReceiptList > 0 ? ccReceiptList[0] : "",
    user_name: userName,
    sender: fromMail,
    subject: subject,
    body: jsonEscape(acutalBody),
    cert_id: randomstring,
    sending_date_time: new Date().toLocaleString(),
    delivered_date_time: new Date().toLocaleString(),
    gateway_name: getWayName,
    attachments: attachmentList,
    email_thread: [],
    smtp_details: smtpDetails,
    conversation_id: conversationId,
  });
}

async function getEmailThread(messageId, authToken1, smtpDetails, conversationId, attachmentList) {
  var roamingSettings = Office.context.roamingSettings;
  // Retrieve the token
  var authToken2 = roamingSettings.get("accessToken");
  var request = new XMLHttpRequest();
  console.log("V222222");
  var url = emailUrl + `?$filter=conversationId eq '${conversationId}'`;
  console.log("Function2");
  console.log(url);
  console.log("getEmailThread");
  request.open("GET", url, false);
  request.setRequestHeader("Content-Type", "application/json");
  request.setRequestHeader("Authorization", "Bearer " + authToken2);
  request.onreadystatechange = async function () {
    console.log("getEmailThread FileInfo----->");
    if (request.readyState === 4) {
      if (request.status === 200) {
        var response = JSON.parse(request.responseText);
        // console.log(response.value);
        var emailThread = await manageEmailThread(
          messageId,
          authToken1,
          smtpDetails,
          conversationId,
          attachmentList,
          response.value
        );
        // getEmailSMTPInfo(messageId, authToken1, response.value, conversationId);
        // getEmailSMTPInfo(messageId, authToken1, emailThread, conversationId);
      } else {
        // Handle the error here
        console.log("FileInfo1error----->", request.status);
      }
    } else {
      console.log("FileInfo2error----->", request.responseText);
    }
  };

  console.log("88json Request ----->");
  // var payloadJson = JSON.stringify(json);
  await request.send();
}

async function getEmailSMTPInfo(messageId, authToken1, conversationId) {
  var roamingSettings = Office.context.roamingSettings;
  // Retrieve the token
  var authToken2 = roamingSettings.get("accessToken");
  var request = new XMLHttpRequest();
  console.log("V222222");
  var url = emailUrl + `/${messageId}/$value`;
  console.log("Function3");
  console.log(url);
  console.log("getEmailSMTPInfo");
  request.open("GET", url, false);
  request.setRequestHeader("Content-Type", "application/json");
  request.setRequestHeader("Authorization", "Bearer " + authToken2);
  request.onreadystatechange = function () {
    console.log("2222111IN FileInfo----->");
    if (request.readyState === 4) {
      if (request.status === 200) {
        var response = request.responseText;
        console.log("afterEmailSend---->>>>");
        // console.log(messageId);
        // console.log(authToken1);
        // console.log(emailThread);
        // console.log(response);
        getAttachments(messageId, authToken1, response, conversationId);
      } else {
        // Handle the error here
        console.log("FileInfo1error----->", request.status);
      }
    } else {
      console.log("FileInfo2error----->", request.responseText);
    }
  };

  console.log("222json Request ----->");
  // var payloadJson = JSON.stringify(json);
  await request.send();
}

// function byteArrayToFile(byteArray1, fileName, mimeType) {
//   console.log("byteArray------->>");
//   console.log(byteArray1);

//   const binaryString = atob(byteArray1);
//   const byteArray = new Uint8Array(binaryString.length);
//   for (let i = 0; i < binaryString.length; i++) {
//     byteArray[i] = binaryString.charCodeAt(i);
//   }
//   console.log(fileName);
//   console.log(mimeType);
//   const blob = new Blob([byteArray], {
//     type: mimeType,
//   });

//   if (navigator.msSaveBlob) {
//     // For Internet Explorer
//     navigator.msSaveBlob(blob, fileName);
//   } else {
//     // For modern browsers
//     const link = document.createElement("a");
//     link.href = URL.createObjectURL(blob);
//     link.download = fileName;
//     link.click();

//     // Clean up the object URL after download
//     URL.revokeObjectURL(link.href);
//   }
// }
function download(content, fileName, contentType) {
  var a = document.createElement("a");
  var file = new Blob([content], {
    type: contentType,
  });
  a.href = URL.createObjectURL(file);
  a.download = fileName;
  a.click();
}

async function afterEmailSend(json) {
  console.log("IN Mail----->afterEmailSend");
  console.log(json);
  var payloadJson1 = JSON.stringify(json);
  download(payloadJson1, "json.txt", "text/plain");
  Object.entries(json).forEach(([key, value]) => {
    console.log(key, value);
  });

  Object.entries(json.attachments).forEach(([key, value]) => {
    console.log(key, value);
  });

  var request = new XMLHttpRequest();
  var url = webHookUrl + "emailInfo";

  request.open("POST", url, false);
  request.setRequestHeader("Content-Type", "application/json");

  console.log("111IN Mail----->");
  request.onreadystatechange = function () {
    console.log("2222111IN Mail----->");
    if (request.readyState === 4) {
      console.log("4444444 Mail----->");
      console.log("responseText----->", request.responseText);
      if (request.status === 200) {
        var response = JSON.parse(request.responseText);
        // Process the response data here
        console.log("response----->", response);
      } else {
        // Handle the error here
        console.log("1error----->", request.status);
      }
    } else {
      console.log("2error----->", request.responseText);
    }
  };

  console.log("66json Request ----->", json);
  var payloadJson = JSON.stringify(json);
  await request.send(payloadJson);
}

function onAddon(event) {
  const item = Office.context.mailbox.item;
  console.log("AKSHAY Mirgal");
  var aRandomstring = generateRandomUniquNo("", 20);
  var strVar =
    `<!doctype html>
<html>
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>Simple Transactional Email</title>
  </head>
  <body style="background-color: #f6f6f6;
  font-family: sans-serif;
  -webkit-font-smoothing: antialiased;
  font-size: 14px;
  margin: 0;
  padding: 0;
  -ms-text-size-adjust: 100%;
  -webkit-text-size-adjust: 100%; ">
    <span class="preheader"></span>
    <table  style = "
      border-collapse: separate;
      mso-table-lspace: 0pt;
      mso-table-rspace: 0pt;
      width: 100%; }
      table td {
        font-family: sans-serif;
        font-size: 14px;
        vertical-align:middle; 
    "role="presentation" border="0" cellpadding="0" cellspacing="0" class="body">
      <tr>
        <td>&nbsp;</td>
        <td style="
          display: block;
          margin: 0 auto !important;
          /* makes it centered */
          max-width: 580px;
          padding: 10px;
          width: 580px; 
        ">
          <div class="
            box-sizing: border-box;
            align-self: center;
            display: block;
            margin: 0 auto;
            width: 100%;
            padding: 10px; 
          ">

            <!-- START CENTERED WHITE CONTAINER -->
            <table role="presentation" style="
              background: #ffffff;
              border-radius: 3px;
              width: 100%; 
            ">

              <!-- START MAIN CONTENT AREA -->

              <tr>
                <td style="    width:100%;
                display: flex;
                vertical-align: middle;
                flex-direction: column;
                gap:2rem;
                background-color: #1E9CC8;">
                  <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td>
                      <div style="display: flex" >
                        <div style="display: flex;
                      flex-direction: column;
                      background-color: #FFFFFF;
                      /* max-height: fit-content; */
                      padding-left: 1rem;
                      padding-right: 1rem;
                      padding-top: 1.0rem;
                      margin: 0.5rem;
                      align-items: center;
                      border-radius: 0.6rem;
                       " >
                      <img
                      src="../assets/logov1.svg"
                      width="50px"
                      height="25px"
                    />
                      <div style="font-weight: 700;font-family: Arial, Helvetica, sans-serif;   font-family: sans-serif;
                      font-size: 12px;
                      font-weight: normal;
                      margin: 0;
                      margin-bottom: 15px; ">authcom.</div>
                    
                      </div>
                      <div class="
                        border-right: 1px solid #E2E7EF;
                        margin-left: 5px;
                        margin-right: 10PX;
                        margin-top: 5px;
                        margin-bottom: 5px;
                      ">
                        
                      </div>
                    </td>

                    <td>
                      <div style="
                        margin-top: 1.2rem;
                        font-size: 0.3rem;
                        font-family: Arial, Helvetica, sans-serif;
                        font-weight: 600;
                        color:#FFFFFF
                
                     ">
                        <div class="" style="text-transform: uppercase;  font-family: sans-serif;
                        font-size: 12px;
                        margin: 0;
                        margin-bottom: 15px; ">Authenticated Receipt</div>
                        <div style="font-weight: 400;  font-family: sans-serif;
                        font-size: 12px;
                        font-weight: normal;
                        margin: 0;
                        margin-bottom: 15px; ">Evidence of Authenticated Email Transaction</div>
                        </div>
                    
                    </td>
                    <td>
                      <div style="
                        font-family: Arial, Helvetica, sans-serif;
                  display: inline-block;
                  padding: 0.25rem 0.5rem;
                  border-radius: 1rem;
                  background-color: #1E9CC8;
                  color: #FFFFFF;
                  font-size: 0.5rem;
                  margin-right: 0.2rem;
                  border: 1px solid #FFFFFF;
                "><p1>Unique ID : ` +
    aRandomstring +
    `</p1></div>
                    </td>
                  </tr>
             
                  </table>
                  <tr>
                  <td class="wrapper">
                    <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                      <div style="display: flex;flex-direction: column; background-color: #FFFFFF; padding-left: 1rem; padding-right: 1rem; padding-top: 1.0rem; padding-bottom: 1.0rem; margin: 0.5rem; border-radius: 0.5rem;">
                      <label for="user-message" class="col-lg-1 control-label">Enter your message below : 
                      </label>
                      <div style="background-color: #EBEBEB;  padding-left: 1.0rem;padding-right: 1.0rem;padding-bottom: 1.0rem; margin-top:1.0rem;"><p></br></br></br></p></div>
                      </dev></tr></table></td></tr></table><div class="footer"><table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="
                  padding-top: 1.0rem;
                   " >
                    Confidential: This certificate has been issued upon request and with the express consent of the sender,
                     by a secure and confidential system. This certificate has been assigned, in the records of
                     the signatory operator, a unique identifier.</a>.
                  </td>
                </tr>
                <img src="${webHookUrl}read-receipt.png/${aRandomstring}" alt="Read Receipt">
               </div>
                </td>
              </tr>
          </div>
        </td>
      </tr>
    </table>
  </body>
</html>
`;

  console.log("111111--------->>>1111-------->>>");
  // Check if a default Outlook signature is already configured.
  item.isClientSignatureEnabledAsync(
    {
      asyncContext: event,
    },
    (result) => {
      console.log("------->>>");
      console.log(result);
      console.log(item.body);
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("111111");
        console.log(result.error.message);
        return;
      }

      item.body.setSignatureAsync(
        strVar,
        {
          asyncContext: event,
          coercionType: Office.CoercionType.Html,
        },
        addSignatureCallback
      );
    }
  );
}

function addSignatureCallback(result) {
  console.log("AKSHAY Mirgalv2");
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.log(result.error.message);
    return;
  }
  console.log("Successfully added signature.");
  result.asyncContext.completed();
}

function validateEmail(emailList) {
  let resList = [];
  if (emailList.length == 0) {
    resList.push(false);
  }
  for (let i = 0; i < emailList.length; i++) {
    if (
      String(emailList[i])
        .toLowerCase()
        .match(
          // eslint-disable-next-line no-useless-escape
          /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
        ) &&
      /@iauro\.com$/.test(emailList[i])
    ) {
      resList.push(true);
    } else {
      resList.push(false);
    }
  }

  console.log(resList);
  return resList.every((item) => item);
  // return false;
}

function onMessageSendHandler(event) {
  let toRecipients, ccRecipients, bccRecipients;

  if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
    toRecipients = item.requiredAttendees;
    ccRecipients = item.optionalAttendees;
  } else {
    toRecipients = item.to;
    ccRecipients = item.cc;
    bccRecipients = item.bcc;
  }

  const toPromise = new Promise((resolve, reject) => {
    toRecipients.getAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        console.log("111111---->");
        toReceiptList = asyncResult.value;
        console.log(toReceiptList[0].emailAddress);
        resolve();
      } else {
        reject();
      }
    });
  });
  if (toRecipients) {
    PromiseList.push(toPromise);
  }

  if (bccRecipients) {
    const bccPromise = new Promise((resolve, reject) => {
      bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status === "succeeded") {
          console.log("222---->");
          bccReceiptList = asyncResult.value;
          resolve();
        } else {
          reject();
        }
      });
    });
    if (bccRecipients) {
      PromiseList.push(bccPromise);
    }
  }

  const ccPromise = new Promise((resolve, reject) => {
    ccRecipients.getAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        console.log("333---->");
        ccReceiptList = asyncResult.value;
        // console.log(ccReceiptList[0].emailAddress);
        resolve();
      } else {
        reject();
      }
    });
  });

  if (ccRecipients) {
    PromiseList.push(ccPromise);
  }

  const fromMailPromise = new Promise((resolve, reject) => {
    item.from.getAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        console.log("4444---->");
        fromMail = asyncResult.value.emailAddress;
        userName = asyncResult.value.displayName;
        console.log(fromMail);
        resolve();
      } else {
        reject();
      }
    });
  });
  PromiseList.push(fromMailPromise);

  const subjectPromise = new Promise((resolve, reject) => {
    console.log("subjectPromise : 000");
    item.subject.getAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        console.log("5555---->");
        subject = asyncResult.value;
        // console.log(subject);
        resolve();
      } else {
        reject();
      }
    });
  });
  PromiseList.push(subjectPromise);

  const bodyPromise = new Promise((resolve, reject) => {
    item.body.getAsync("html", function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        console.log("7777---->");
        body = asyncResult.value;
        // console.log(body);
        resolve();
      } else {
        reject();
      }
    });
  });
  PromiseList.push(bodyPromise);

  Promise.all(PromiseList)
    .then((result) => {
      console.log("Promise --------------->>>>>>>>>");
      console.log(fromMail);
      console.log("111Promise --------------->>>>>>>>>");
      console.log(toReceiptList);
      console.log(bccReceiptList);
      console.log(ccReceiptList);
      console.log(subject);
      console.log(attachmentID);
      console.log("attachmentID-------------------------------->");
      randomstring = body.split(`Unique ID : `).pop().split(`</div>`)[0];
      console.log(randomstring);
      acutalBody = body
        .split(
          `<div style="padding-left: 1rem; padding-right: 1rem; padding-bottom: 1rem; margin-top: 1rem; background-color: rgb(235, 235, 235);">`
        )
        .pop()
        .split(`</div>`)[0]
        .replaceAll("<br>", "");
      acutalBody = acutalBody.replaceAll("</p><p>", "\n");
      acutalBody = acutalBody.replaceAll("</p>", "");
      acutalBody = acutalBody.replaceAll("<p>", "");
      acutalBody = acutalBody.replaceAll("\n", "");
      console.log("AAAAAAAAAAAAAAAA-------------------->>>");
      console.log(acutalBody);
      // sendMail();
      // showDialog();
      // generateFileInfo(item);
      // attachmentCode();
      // getFileInfo("", "", "");
      if (toReceiptList.length > 0 && fromMail != "" && acutalBody != "" && subject != "") {
        console.log("1111------>");
        console.log("BBBBBBBBBBBB-------------------->>>");
        var roamingSettings = Office.context.roamingSettings;
        // Retrieve the token
        var token = roamingSettings.get("accessToken");

        if (token) {
          getFileInfo(token);
        } else {
          console.log("Token not found in roaming settings.");
        }

        event.completed({
          allowEvent: true,
        });
      } else {
        console.log("22222------>");
        let message = "Failed to get body text";
        console.error(message);
        event.completed({
          allowEvent: true,
          errorMessage: message,
        });
      }
    })
    .catch((error) => {
      throw new Error(error);
    });
}

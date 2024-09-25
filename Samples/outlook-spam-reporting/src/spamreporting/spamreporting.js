/*
 * Copyright (c) Riccio78
 * Licensed under the MIT license.
 */

/* global Office */
/* global console */

// Authentication for MS Graph API
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { ClientSecretCredential } from "@azure/identity";
import {
  clientId,
  clientSecret,
  tenantId,
  mailSubjectTestMail,
  mailSubjectSpamMail,
  mailAddress,
  titleCompletedTestMail,
  textCompletedTestMail,
  titleCompletedSpamMail,
  textCompletedSpamMail,
  usedPhishingTestHeader,
} from "../config/vars.js";

// Ensures the Office.js library is loaded.
Office.onReady();

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  let isPhishingTestMail = false;

  //Checking headers if customer header indicated phishing test message
  Office.context.mailbox.item
    .getAllInternetHeadersAsync()
    .then((headers) => {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i].name.toLowerCase() === usedPhishingTestHeader.toLowerCase()) {
          // header found, test mail was sent by IT deparment
          isPhishingTestMail = true;
          break;
        }
      }
    })
    .catch((error) => {
      console.error("Fehler beim Abrufen der Header:", error);
    });

  /* 	Office.context.mailbox.item.getAsFileAsync().then(file => {
			// The file is now available
			console.log(file);
			// Do something with the file, e.g., save it, upload it, or analyze it
		}); 
	*/

  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync({ asyncContext: event }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
      return;
    }

    var emailAttachment = asyncResult.value;
    // var reader = new FileReader();
    // reader.onload = function (e) {
    //   // working with file at this point (e.target.result)
    //   console.log(e.target.result);
    // };
    // reader.readAsText(emailAttachment);

    // Get the user's responses to the options and text box in the preprocessing dialog.
    const spamReportingEvent = asyncResult.asyncContext;
    const reportedOptions = spamReportingEvent.options;
    const additionalInfo = spamReportingEvent.freeText;

    // Run additional processing operations here.

    // sending mails without user interaction requires server side
    // technology like Microsoft Graph API.
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      // The client credentials flow requires that you request the
      // /.default scope, and pre-configure your permissions on the
      // app registration in Azure. An administrator must grant consent
      // to those permissions beforehand.
      scopes: ["https://graph.microsoft.com/.default"],
    });

    const clientMSGraph = Client.initWithMiddleware({ authProvider: authProvider });

    // unsolicited message, configure subject and add email as attachment.
    async function sendEmailWithAttachment() {
      const message = {
        subject: mailSubjectSpamMail,
        body: {
          contentType: "Text",
          content: reportedOptions + "\n\n" + additionalInfo,
        },
        toRecipients: [
          {
            emailAddress: {
              address: mailAddress,
            },
          },
        ],
        attachments: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: "reportedEmail.eml",
            contentBytes: emailAttachment,
            //name: emailAttachment,
            //contentBytes: "BASE64_ENCODED_EMAIL_CONTENT",
          },
        ],
      };

      await clientMSGraph.api("/me/sendMail").post({ message });
      console.log("E-Mail gesendet");
    }

    // Test message, configure subject. Do not add attachment
    // sending this message enables statistical evaluation.
    async function sendEmail() {
      const message = {
        subject: mailSubjectTestMail,
        body: {
          contentType: "Text",
          content: reportedOptions + "\n\n" + additionalInfo,
        },
        toRecipients: [
          {
            emailAddress: {
              address: mailAddress,
            },
          },
        ],
      };

      await clientMSGraph.api("/me/sendMail").post({ message });
      console.log("E-Mail gesendet");
    }

    if (isPhishingTestMail) {
      sendEmail().catch(console.error);
    } else {
      sendEmailWithAttachment().catch(console.error);
    }

    /**
     * Signals that the spam-reporting event has completed processing.
     * It then moves the reported message to the Junk Email folder of the mailbox,
     * then shows a post-processing dialog to the user.
     * If an error occurs while the message is being processed,
     * the `onErrorDeleteItem` property determines whether the message will be deleted.
     */
    const event = asyncResult.asyncContext;
    if (isPhishingTestMail) {
      event.completed({
        onErrorDeleteItem: true,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.DeletedItemsFolder,
        showPostProcessingDialog: {
          title: titleCompletedTestMail,
          description: textCompletedTestMail,
        },
      });
    } else {
      event.completed({
        onErrorDeleteItem: true,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: titleCompletedSpamMail,
          description: textCompletedSpamMail,
        },
      });
    }
  });
}

/**
 * IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name
 * specified in the manifest to its JavaScript counterpart.
 */
Office.actions.associate("onSpamReport", onSpamReport);

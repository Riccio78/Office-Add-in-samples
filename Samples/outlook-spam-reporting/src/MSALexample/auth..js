const config = {
    auth: {
        clientId: "your-client-id",
        authority: "https://login.microsoftonline.com/your-tenant-id",
        clientSecret: "your-client-secret"
    }
};

module.exports = config;

const { ConfidentialClientApplication } = require("@azure/msal-node");
const config = require("./config");

const cca = new ConfidentialClientApplication(config);

const clientCredentialRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
};

cca.acquireTokenByClientCredential(clientCredentialRequest).then((response) => {
    console.log("Access Token:", response.accessToken);
    // Hier kannst du den Access Token verwenden, um Anfragen an Microsoft Graph zu senden
}).catch((error) => {
    console.error("Error acquiring token:", error);
});


const { ConfidentialClientApplication } = require("@azure/msal-node");
const axios = require("axios");
const config = require("./config");

const cca = new ConfidentialClientApplication(config);

const clientCredentialRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
};

cca.acquireTokenByClientCredential(clientCredentialRequest).then((response) => {
    const accessToken = response.accessToken;

    // API-Aufruf an Microsoft Graph
    axios.get("https://graph.microsoft.com/v1.0/users", {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    }).then((res) => {
        console.log("Benutzerliste:", res.data);
    }).catch((error) => {
        console.error("Fehler beim Abrufen der Benutzerliste:", error);
    });
}).catch((error) => {
    console.error("Fehler beim Abrufen des Tokens:", error);
});


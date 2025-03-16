const msal  = require("@azure/msal-node");
const { PublicClientApplication, CryptoProvider, Configuration, ConfidentialClientApplication } = msal
const graph = require("@microsoft/microsoft-graph-client")

/*
employee is available if he working on that day at that time.
If he is status set on do not disturbe
If he is not on vecation

*/
const tenantId = process.env.AZURE_TENANT_ID //"283d6102-3937-4c3b-8381-cbec140bdef8";
const clientId = process.env.AZURE_CLIENT_ID // "f53a37a2-c624-4c76-9c26-083c8109678c";


console.log(tenantId)
console.log(clientId)



const appMsalConfig = {
    auth: {
        clientId: clientId,
        authority: `https://auth.microsoftonline.com/${tenantId}`,
        clientSecret: // TODO
    },
    system: {
        loggerOptions: {
            loggerCallback(logLevel, message, containsPii) {
                if(!containsPii) console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
}

let confidentialClient = new ConfidentialClientApplication(appMsalConfig)



module.exports = {
    setup: () => {
        console.log("Setup confidential client application...")
        confidentialClient = new ConfidentialClientApplication(appMsalConfig)
    },
    getUserMailboxSettings: async function() {
        const client = getAuthenticationClient(confidentialClient)

        console.log("getUserMailboxSettings")

        const res = await client
            .api("/users/530da9b6-18f9-4f1d-a66d-bc502ef6c169/mailboxSettings")
            .header("Prefer", `outlook.timezone="Etc/GMT"`)
            .get()
    },
    getUserP
}


/**
 *
 * @param msalClient
 */
function getAuthenticationClient(msalClient) {
    if (!msalClient) {
        throw new Error("Missing authentication client")
    }

    return graph.Client.init({
        authProvider: async (done) => {

            try {
                const config = {
                    authority: `https://login.microsoftonline.com/${tenantId}`,
                    correlationId: "test",
                    scopes: [
                        "https://graph.microsoft.com/.default"
                    ]
                }

                const result = await confidentialClient.acquireTokenByClientCredential(config)

                done(null, result.accessToken)

            } catch (err) {
                console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
                done(err, null)
            }
        }
    })
}
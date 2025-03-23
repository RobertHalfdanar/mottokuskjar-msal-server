const msal  = require("@azure/msal-node");
const { PublicClientApplication, CryptoProvider, Configuration, ConfidentialClientApplication } = msal
const graph = require("@microsoft/microsoft-graph-client")

/*
employee is available if he working on that day at that time.
If he is status set on do not disturbe
If he is not on vecation

*/
const tenantId = process.env.AZURE_TENANT_ID
const clientId = process.env.AZURE_CLIENT_ID
const clientSecret = process.env.AZURE_CLIENT_SECRET


const appMsalConfig = {
    auth: {
        clientId: clientId,
        authority: `https://auth.microsoftonline.com/${tenantId}`,
        clientSecret: clientSecret
    },
}

let confidentialClient = new ConfidentialClientApplication(appMsalConfig)



module.exports = {
    setup: () => {
        console.log("Setup confidential client application...")
        confidentialClient = new ConfidentialClientApplication(appMsalConfig)
    },

    /**
     *
     * @returns {Promise<EmployeesRespond>}
     */
    getUsersInGroup: async function () {
        const client = getAuthenticationClient(confidentialClient)

        return await client
            .api("/groups/6f44b283-d134-47f8-90fe-d6ae931b418d/members/microsoft.graph.user")
            .get()
    },
    /**
     * @param {string} id
      * @returns {Promise<MailboxSettings>}
     */
    getUserMailboxSettings: async function(id) {
        const client = getAuthenticationClient(confidentialClient)

        return await client
            .api(`/users/${id}/mailboxSettings`)
            .header("Prefer", `outlook.timezone="Etc/GMT"`)
            .get()

    },

    /**
     * @param {string} id
     * @returns {Promise<Presence>}
     */
    getUserPresence: async function(id) {
        const client = getAuthenticationClient(confidentialClient)

        return await client
            .api(`/users/${id}/presence`)
            .header("Prefer", `outlook.timezone="Etc/GMT"`)
            .get()

    }
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
                done(err, null)
            }
        }
    })
}

/** @typedef {object} MailboxSettings
 * @property {string} @odata.context
 * @property {string} archiveFolder
 * @property {string} timeZone
 * @property {string} delegateMeetingMessageDeliveryOptions
 * @property {string} dateFormat
 * @property {string} timeFormat
 * @property {string} userPurpose
 * @property {object} automaticRepliesSetting
 * @property {string} automaticRepliesSetting.status
 * @property {string} automaticRepliesSetting.externalAudience
 * @property {string} automaticRepliesSetting.internalReplyMessage
 * @property {string} automaticRepliesSetting.externalReplyMessage
 * @property {object} automaticRepliesSetting.scheduledStartDateTime
 * @property {string} automaticRepliesSetting.scheduledStartDateTime.dateTime
 * @property {string} automaticRepliesSetting.scheduledStartDateTime.timeZone
 * @property {object} automaticRepliesSetting.scheduledEndDateTime
 * @property {string} automaticRepliesSetting.scheduledEndDateTime.dateTime
 * @property {string} automaticRepliesSetting.scheduledEndDateTime.timeZone
 * @property {object} language
 * @property {string} language.locale
 * @property {string} language.displayName
 * @property {object} workingHours
 * @property {string[]} workingHours.daysOfWeek
 * @property {string} workingHours.startTime
 * @property {string} workingHours.endTime
 * @property {object} workingHours.timeZone
 * @property {string} workingHours.timeZone.name
 */


/** @typedef {object} Presence
 * @property {string} @odata.context
 * @property {string} id
 * @property {string} availability
 * @property {string} activity
 * @property {null} statusMessage
 */


/** @typedef {object} Employee
 * @property {} businessPhones
 * @property {string} displayName
 * @property {string} givenName
 * @property {null} jobTitle
 * @property {string} mail
 * @property {null|string} mobilePhone
 * @property {null} officeLocation
 * @property {null|string} preferredLanguage
 * @property {string} surname
 * @property {string} userPrincipalName
 * @property {string} id
 */


/** @typedef {object} EmployeesRespond
 * @property {string} @odata.context
 * @property {Employee[]} value
 */

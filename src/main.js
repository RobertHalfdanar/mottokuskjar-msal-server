
const {
	CommunicationIdentityClient,
} = require("@azure/communication-identity");
const msal  = require("@azure/msal-node");
const { PublicClientApplication, CryptoProvider, Configuration, ConfidentialClientApplication } = msal

const graph = require("./graph")

const express = require("express");

// You will need to set environment variables in .env
const SERVER_PORT = process.env.PORT || 3000
const SERVER_HOST = process.env.HOST || "http://localhost:3000"

const clientId = process.env.AZURE_CLIENT_ID // "f53a37a2-c624-4c76-9c26-083c8109678c";
const tenantId = process.env.AZURE_TENANT_ID //"283d6102-3937-4c3b-8381-cbec140bdef8";


const configurations = {
	"graph": {
		"scopes": [
			"Group.Read.All",
			"user.read",
			"User.Read.All",
		],
		"redirectUri": `${SERVER_HOST}/graph/redirect` // "http://localhost:3000/graph/redirect",
	},
	"teams": {
		"scopes": [
			"https://auth.msft.communication.azure.com/Teams.ManageCalls",
			"https://auth.msft.communication.azure.com/Teams.ManageChats",
		],
		"redirectUri": `${SERVER_HOST}/teams/redirect`,  // "http://localhost:3000/teams/redirect",
	}
}


// Quickstart code goes here
// Create configuration object that will be passed to MSAL instance on creation.dev
/** @type Configuration */
const msalConfig = {
	auth: {
		clientId: clientId,
		authority: `https://login.microsoftonline.com/${tenantId}`,
		redirectUri: configurations.graph.redirectUri,
	},
};

// Create an instance of PublicClientApplication
const pca = new PublicClientApplication(msalConfig);
const provider = new CryptoProvider();

const app = express();

let pkceVerifier = "";

app.get("/graph", async (req, res) => {
	// Generate PKCE Codes before starting the authorization flow
	const { verifier, challenge } = await provider.generatePkceCodes();
	pkceVerifier = verifier;

	const authCodeUrlParameters = {
		scopes: configurations.graph.scopes,
		redirectUri: configurations.graph.redirectUri,
		codeChallenge: challenge,
		codeChallengeMethod: "S256",
	};
	// Get url to sign user in and consent to scopes needed for application
	pca
		.getAuthCodeUrl(authCodeUrlParameters)
		.then((response) => {
			res.redirect(response);
		})
		.catch((error) => console.log(JSON.stringify(error)));
});

app.get("/teams", async (req, res) => {
	// Generate PKCE Codes before starting the authorization flow
	const { verifier, challenge } = await provider.generatePkceCodes();
	pkceVerifier = verifier;

	const authCodeUrlParameters = {
		scopes: configurations.teams.scopes,
		redirectUri: configurations.teams.redirectUri,
		codeChallenge: challenge,
		codeChallengeMethod: "S256",
	};
	// Get url to sign user in and consent to scopes needed for application
	pca
		.getAuthCodeUrl(authCodeUrlParameters)
		.then((response) => {
			res.redirect(response);
		})
		.catch((error) => console.log(JSON.stringify(error)));
});

app.get("/graph/redirect", async (req, res) => {
	// Create request parameters object for acquiring the AAD token and object ID of a Teams user

	const tokenRequest = {
		code: req.query.code,
		scopes: configurations.graph.scopes,
		redirectUri: configurations.graph.redirectUri,
		codeVerifier: pkceVerifier,
	};
	console.log(tokenRequest);

	// Retrieve the AAD token and object ID of a Teams user
	pca
		.acquireTokenByCode(tokenRequest)
		.then(async (response) => {
			console.log("Response:", response);
			let token = response.accessToken;
			let expiresOn = response.expiresOn.toISOString();

			res.redirect(`/token?token=${token}&expiresOn=${expiresOn}`);
		})
		.catch((error) => {
			console.log(error);
			res.status(500).send(error);
		});
});

app.get("/teams/redirect", async (req, res) => {
	// Create request parameters object for acquiring the AAD token and object ID of a Teams user

	const tokenRequest = {
		code: req.query.code,
		scopes: configurations.teams.scopes,
		redirectUri: configurations.teams.redirectUri,
		codeVerifier: pkceVerifier,
	};
	console.log(tokenRequest);

	// Retrieve the AAD token and object ID of a Teams user
	pca
		.acquireTokenByCode(tokenRequest)
		.then(async (response) => {
			console.log("Response:", response);
			let token = response.accessToken;
			let expiresOn = response.expiresOn.toISOString();

			console.log("Expires on: " + response.expiresOn)

			let userObjectId = response.uniqueId;

			const connectionString = "endpoint=https://mottokuskjar.europe.communication.azure.com/;accesskey=8GKKJN26NZMHi2R35p7D4xXwSPmTF7VUpYwoONrWSvd4xmm2aaGhJQQJ99AKACULyCpY44QIAAAAAZCSbcPB"
			const identityClient = new CommunicationIdentityClient(connectionString);

			try {
				const communicationAccessToken = await identityClient.getTokenForTeamsUser({
					teamsUserAadToken: token,
					clientId: clientId,
					userObjectId: userObjectId,
				});

				token = communicationAccessToken.token
				expiresOn = communicationAccessToken.expiresOn.toISOString()

			} catch (error) {
				console.log(error);
			}

			res.redirect(`/token?token=${token}&expiresOn=${expiresOn}`);
		})
		.catch((error) => {
			console.log(error);
			res.status(500).send(error);
		});
});

app.get("/token", async (req, res) => {

	res.send("Hello World!");
})


app.get("/employees", async (req, res) => {


	const employees = await graph.getUsersInGroup()




	/**
	 *
	 * @param {Presence} presence
	 */
	function evaluateUserPresence(presence) {
		return presence.availability !== "DoNotDisturb";
	}

	/**
	 *
	 * @param {MailboxSettings} mailboxSettings
	 */
	function evaluateUserMailboxSetting(mailboxSettings) {
		const today = new Date()

		const weekdays = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
		const weekday = weekdays[today.getDay()]


		console.log(mailboxSettings)

		if (!mailboxSettings.workingHours.daysOfWeek.includes(weekday)) return false

		const startTimeHours = mailboxSettings.workingHours.startTime.slice(0, 2)
		const startTimeMinutes = mailboxSettings.workingHours.startTime.slice(3, 5)

		const endTimeHours = parseInt(mailboxSettings.workingHours.endTime.slice(0, 2))
		const endTimeMinutes = parseInt(mailboxSettings.workingHours.endTime.slice(3, 5))

		const startDate = new Date()
		startDate.setHours(startTimeHours)
		startDate.setMinutes(startTimeMinutes)

		const endDate = new Date()
		endDate.setHours(endTimeHours)
		endDate.setMinutes(endTimeMinutes)

		return startDate <= today && today <= endDate
	}


	/**
	 * @param {MailboxSettings} mailboxSettings
	 */
	function evaluateUserVacation(mailboxSettings) {
		if (mailboxSettings.automaticRepliesSetting.status === "disabled") return true

		console.log(mailboxSettings)

		const startDate = new Date(mailboxSettings.automaticRepliesSetting.scheduledStartDateTime.dateTime)
		const endDate = new Date(mailboxSettings.automaticRepliesSetting.scheduledEndDateTime.dateTime)

		const today = new Date()

		return !(startDate <= today && today <= endDate)
	}

	const employeesState = {}


	for (const employee of employees.value) {

		//if (employee.displayName != "Róbert Hálfdanarson") continue

		const mailboxSetting = await graph.getUserMailboxSettings(employee.id)
		const presence = await graph.getUserPresence(employee.id)

		employeesState[employee.displayName] = evaluateUserMailboxSetting(mailboxSetting) && evaluateUserVacation(mailboxSetting) && evaluateUserPresence(presence)
	}

	employees.value.sort((a, b) => a.displayName.localeCompare(b.displayName))


	res.json({employees: employees.value.map(emp => ({
		name: emp.displayName,
		id: emp.id,
		isActive: employeesState[emp.displayName],
	}))})

})

app.listen(SERVER_PORT, () => {
		graph.setup()
		console.log(
			`Communication access token application started on ${SERVER_PORT}!`
		)
	}
);

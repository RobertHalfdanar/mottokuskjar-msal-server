const {
	CommunicationIdentityClient,
} = require("@azure/communication-identity");
const { PublicClientApplication, CryptoProvider, Configuration } = require("@azure/msal-node");
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

app.listen(SERVER_PORT, () =>
	console.log(
		`Communication access token application started on ${SERVER_PORT}!`
	)
);

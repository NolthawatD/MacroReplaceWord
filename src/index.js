import * as dotenv from "dotenv";
import express from "express";
import cors from "cors";

import { specs } from "./util/swagger.js";
import swaggerUi from "swagger-ui-express";
import { googleSheetRouter } from "./router/google-sheet.router.js";
import { googleDocRouter } from "./router/google.router.js";
import { convertObjectCredentialToJSONFile } from "./util/convert-json.js";
dotenv.config();

const port = process.env.PORT || 8080;
if (!port) {
	process.exit(1);
}

// #create credential file from .env

async function createFileCredential() {
	const credentials = {
		installed: {
			client_id: process.env.CREDENTIAL_AUTH_CLIENT_ID,
			// project_id: process.env.CREDENTIAL_AUTH_PROJECT_ID,
			// auth_uri: process.env.CREDENTIAL_AUTH_AUTH_URI,
			// token_uri: process.env.CREDENTIAL_AUTH_TOKEN_URI,
			// auth_provider_x509_cert_url: process.env.CREDENTIAL_AUTH_AUTH_PROVIDER_X509_CERT_URL,
			client_secret: process.env.CREDENTIAL_AUTH_CLIENT_SECRET,
			redirect_uris: ["http://localhost"],
		},
	};
	await convertObjectCredentialToJSONFile(credentials, "credentials-auth");
}

async function createFileService() {
	console.log("process.env.SERVICE_PRIVATE_KEY === ", process.env.SERVICE_PRIVATE_KEY)

	const credentials = {
		// type: process.env.SERVICE_TYPE,
		// project_id: process.env.SERVICE_PROJECT_ID,
		// private_key_id: process.env.SERVICE_PRIVATE_KEY_ID,
		private_key: process.env.SERVICE_PRIVATE_KEY,
		client_email: process.env.SERVICE_CLIENT_EMAIL,
		// client_id: process.env.SERVICE_CLIENT_ID,
		// auth_uri: process.env.SERVICE_AUTH_URI,
		// token_uri: process.env.SERVICE_TOKEN_URI,
		// auth_provider_x509_cert_url: process.env.SERVICE_AUTH_PROVIDER_X509_CERT_URL,
		// client_x509_cert_url: process.env.SERVICE_CLIENT_X509_CERT_URL,
		// universe_domain: process.env.SERVICE_UNIVERSE_DOMAIN,
	};
	console.log("%c  === ","color:cyan","  credentials", credentials)

	await convertObjectCredentialToJSONFile(credentials, "credentials-sheet");
}


await createFileCredential();
await createFileService();

// #express

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use("/api-docs", swaggerUi.serve, swaggerUi.setup(specs));

app.use("/api/v1", googleSheetRouter);
app.use("/api/v1", googleDocRouter);

app.listen(port, (req, res) => {
	console.log("http://localhost:8080/api-docs");
});

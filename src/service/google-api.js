import * as fs from "fs/promises";
import path from "path";
import process from "process";
import { authenticate } from "@google-cloud/local-auth";
import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import { createWriteStream, createReadStream } from "fs";
import { deleteFilesInFolderAsync } from "../util/fs-handler.js";
import { convertDocxToPDF } from "./convert-file.js";
import { createArrayObjectToBatchUpdate, createArrayObjectToBatchUpdateDemo1 } from "../util/util-google.js";
import { throwError } from "../util/throw-error.js";
import { STATUS_CODES } from "../util/status-code.js";

// If modifying these scopes, delete token.json.
const SCOPES = [
	"https://www.googleapis.com/auth/documents",
	"https://www.googleapis.com/auth/documents.readonly",
	"https://www.googleapis.com/auth/drive.readonly",
	"https://www.googleapis.com/auth/drive.file", // or "https://www.googleapis.com/auth/drive"
];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), "./src/config/token.json");
const CREDENTIALS_PATH = path.join(process.cwd(), "./src/config/credentials-auth.json");
const PATH_FOLDER_KEEP_DOC = `./download/doc/`;

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
	try {
		const content = await fs.readFile(TOKEN_PATH);
		const credentials = JSON.parse(content);
		return google.auth.fromJSON(credentials);
	} catch (err) {
		return null;
	}
}

/**
 * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
	const content = await fs.readFile(CREDENTIALS_PATH);
	const keys = JSON.parse(content);
	const key = keys.installed || keys.web;
	const payload = JSON.stringify({
		type: "authorized_user",
		client_id: key.client_id,
		client_secret: key.client_secret,
		refresh_token: client.credentials.refresh_token,
	});
	await fs.writeFile(TOKEN_PATH, payload);
}

async function authorize() {
	let client = await loadSavedCredentialsIfExist();
	if (client) {
		return client;
	}

	client = await authenticate({
		scopes: SCOPES,
		keyfilePath: CREDENTIALS_PATH,
	});
	if (client.credentials) {
		await saveCredentials(client);
	}
	return client;
}

/**
 * Prints the title of a sample doc:
 * https://docs.google.com/document/d/195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth 2.0 client.
 */
async function printDocTitle(documentId) {
	console.log("%c  === ", "color:cyan", "func printDocTitle documentId", documentId);
	const auth = await authorize();
	try {
		const docs = google.docs({ version: "v1", auth });
		const res = await docs.documents.get({
			documentId: documentId,
		});

		// console.log(`Response : `, res);
		// console.log(`The title of the document is: ${res.data.title}`);
		return res.data.title;
	} catch (error) {
		console.error("Error not found document:", error);
		throw error;
	}
}

async function copyGoogleDoc(sourceDocId, copyTitle) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const copyResponse = await drive.files.copy({
			fileId: sourceDocId,
			requestBody: {
				name: copyTitle || "Copy of Document",
			},
		});

		const fileId = copyResponse.data.id;

		await drive.permissions.create({
			fileId: fileId,
			requestBody: {
				role: "reader",
				type: "anyone",
			},
		});

		return fileId;
	} catch (error) {
		console.error("Error copying document:", error);
		throw error;
	}
}

async function deleteGoogleDoc(docId) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		await drive.files.delete({
			fileId: docId,
		});

		console.log(`Document deleted successfully. Document ID: ${docId}`);
	} catch (error) {
		console.error("Error deleting document:", error);
		throw error;
	}
}

async function exportDocAsWord(docId, nameFile) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });
	const docs = google.docs({ version: "v1", auth });

	try {
		const res = await docs.documents.get({
			documentId: docId,
		});

		const fileId = res.data.documentId;
		const exportOptions = {
			mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		};

		const exportRes = await drive.files.export(
			{
				fileId: fileId,
				mimeType: exportOptions.mimeType,
			},
			{ responseType: "stream" }
		);

		const outputPath = path.join(process.cwd(), `${PATH_FOLDER_KEEP_DOC}${nameFile}.docx`);
		const fileStream = createWriteStream(outputPath);
		exportRes.data.pipe(fileStream);

		return new Promise((resolve, reject) => {
			fileStream.on("finish", resolve);
			fileStream.on("error", reject);
		});
	} catch (error) {
		console.error("Error exporting the document:", error);
	}
}

async function exportGoogleDocAsPDF(docId, pathSaveFilePDF) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const exportOptions = {
			mimeType: "application/pdf",
		};

		const exportRes = await drive.files.export(
			{
				fileId: docId,
				mimeType: exportOptions.mimeType,
			},
			{ responseType: "stream" }
		);

		const outputPath = path.join(process.cwd(), pathSaveFilePDF);
		const fileStream = createWriteStream(outputPath);
		exportRes.data.pipe(fileStream);

		return await new Promise((resolve, reject) => {
			fileStream.on("finish", resolve);
			fileStream.on("error", reject);
		});
	} catch (error) {
		console.error("Error exporting document as PDF:", error);
		throw error;
	}
}

async function writeDocBatchUpdate(docId, requests) {
	const auth = await authorize();
	const docs = google.docs({ version: "v1", auth });
	try {
		const response = await docs.documents.batchUpdate({
			documentId: docId,
			resource: {
				requests,
			},
		});
		// console.log("%c  === ", "color:cyan", "  response", response);
	} catch (error) {
		console.error("Error writing to the document:", error);
	}
}

/***
 *
 *    === SHEET ===
 *
 */

/**
 * Copies a Google Sheet.
 * @param {string} sourceSheetId The ID of the source Google Sheet.
 * @param {string} copyTitle The title of the copied Google Sheet.
 * @returns {Promise<string>} The ID of the copied Google Sheet.
 */
async function copyGoogleSheet(sourceSheetId, copyTitle) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const copyResponse = await drive.files.copy({
			fileId: sourceSheetId,
			requestBody: {
				name: copyTitle || "Test Copy of Sheet",
			},
		});

		console.log("%c  === ", "color:cyan", "  copyResponse", copyResponse.data.id);

		return copyResponse.data.id;
	} catch (error) {
		console.error("Error copying Google Sheet:", error);
		throw error;
	}
}

/**
 * Creates a new spreadsheet.
 * @param {string} title The title of the new spreadsheet.
 * @returns {Promise<string>} The ID of the newly created spreadsheet.
 */
async function createNewSpreadsheet(title) {
	const auth = await authorize();
	const sheets = google.sheets({ version: "v4", auth });
	const drive = google.drive({ version: "v3", auth });

	try {
		const spreadsheet = await sheets.spreadsheets.create({
			resource: {
				properties: {
					title: title,
				},
			},
			fields: "spreadsheetId",
		});

		const fileId = spreadsheet.data.spreadsheetId;

		// Set sharing permissions to public
		await drive.permissions.create({
			fileId: fileId,
			requestBody: {
				role: "reader",
				type: "anyone",
			},
		});
		return fileId;
	} catch (error) {
		console.error("Error creating new spreadsheet:", error);
		throw error;
	}
}

/**
 * Writes data to a specific range in a spreadsheet.
 * @param {string} spreadsheetId The ID of the spreadsheet.
 * @param {string} range The range to write data to (e.g., "Sheet1!A1").
 * @param {Array<Array<string>>} values The values to write to the range.
 * @returns {Promise<void>}
 */
async function writeDataToSheet(spreadsheetId, range, values) {
	const auth = await authorize();
	const sheets = google.sheets({ version: "v4", auth });

	try {
		const resource = {
			values: values,
		};

		const writetDataToSpreadSheet = await sheets.spreadsheets.values.update({
			spreadsheetId: spreadsheetId,
			range: range,
			valueInputOption: "RAW",
			resource: resource,
		});

		return writetDataToSpreadSheet;
	} catch (error) {
		console.error("Error writing data to sheet:", error);
		throw error;
	}
}

async function createNewSpreadsheetWithData(title, sheetData, range) {
	try {
		// Create a new spreadsheet
		const spreadsheetId = await createNewSpreadsheet(title);
		const writeData = await writeDataToSheet(spreadsheetId, range, sheetData);
		return { spreadsheetId, writeData };
	} catch (error) {
		console.error(error);
		throw error;
	}
}

/**
 *
 * 		=== DOWLOAD ===
 *
 */

async function downloadGoogleSheetsXlsx(fileId, fileName) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const exportOptions = {
			mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // XLSX format
		};

		const exportRes = await drive.files.export(
			{
				fileId: fileId,
				mimeType: exportOptions.mimeType,
			},
			{ responseType: "stream" }
		);

		const outputPath = `./${fileName}.xlsx`;
		const fileStream = createWriteStream(outputPath);
		exportRes.data.pipe(fileStream);

		return new Promise((resolve, reject) => {
			fileStream.on("finish", resolve);
			fileStream.on("error", reject);
		});
	} catch (error) {
		console.error("Error downloading Google Sheets file:", error);
		throw error;
	}
}

async function downloadGoogleSheetsXlsxPipe(fileId, res, fileName) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const exportOptions = {
			mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // XLSX format
		};

		const exportRes = await drive.files.export(
			{
				fileId: fileId,
				mimeType: exportOptions.mimeType,
			},
			{ responseType: "stream" }
		);

		// Set response headers
		res.setHeader("Content-Type", exportOptions.mimeType);
		res.setHeader("Content-Disposition", `attachment; filename="${fileName}.xlsx"`);

		// Stream the file contents to the response object
		exportRes.data.pipe(res);
	} catch (error) {
		console.error("Error downloading Google Sheets file:", error);
		throw error;
	}
}

/**
 * Downloads a Google Docs file in DOCX format and sends it directly to the user.
 * @param {string} fileId The ID of the Google Docs file.
 * @param {Object} res The response object to send the file to.
 * @param {string} fileName The name to send the downloaded file as.
 * @returns {Promise<void>}
 */
async function downloadGoogleDocsDocxPipe(fileId, res, fileName) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const exportOptions = {
			mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // DOCX format
		};

		const exportRes = await drive.files.export(
			{
				fileId: fileId,
				mimeType: exportOptions.mimeType,
			},
			{ responseType: "stream" }
		);

		// Set response headers
		res.setHeader("Content-Type", exportOptions.mimeType);
		res.setHeader("Content-Disposition", `attachment; filename="${fileName}.docx"`);

		// Stream the file contents to the response object
		exportRes.data.pipe(res);

		await drive.files.delete({
			fileId: fileId,
		});
	} catch (error) {
		console.error("Error downloading Google Docs file:", error);
		throw error;
	}
}

async function uploadFileToDrive(filePath, folderId, fileName) {
	console.log("folderId in uploadFileToDrive function ===== >", folderId);

	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });

	try {
		const fileMetadata = {
			name: fileName,
			parents: [folderId], // Specify the parent folder ID
		};

		const media = {
			mimeType: "application/octet-stream",
			body: createReadStream(filePath),
		};

		const response = await drive.files.create({
			resource: fileMetadata,
			media: media,
			fields: "id",
		});

		return response.data.id;
	} catch (error) {
		console.error("Error uploading file to Google Drive:", error);
		throw error;
	}
}

async function rebuildDocs(docId, dataBatchArr) {
	try {
		// create new folder in download
		// and send to export to keep data
		// and return pwd to response to dowload

		// #remove file all in folder
		await deleteFilesInFolderAsync(PATH_FOLDER_KEEP_DOC);

		await Promise.all(
			dataBatchArr.map(async (value, i) => {
				// console.log("%c  === ", "color:cyan", "  docId", docId);
				// console.log("%c  === ", "color:cyan", "  value", value);
				const nameCopy = "title_demo_" + (i + 1);
				const copyDocId = await copyGoogleDoc(docId, nameCopy);
				const dataToReplace = await createArrayObjectToBatchUpdate(value);
				// console.log("%c  === ", "color:cyan", "  dataToReplace", dataToReplace);

				await writeDocBatchUpdate(copyDocId, dataToReplace);
				await exportDocAsWord(copyDocId, nameCopy);
				// await exportGoogleDocAsPDF(copyDocId, nameCopy);
				await deleteGoogleDoc(copyDocId);
			})
		);
	} catch (error) {
		console.error("Error copying document:", error);
		throw error;
	}
}

async function rebuildDocsDemo1(docId, dataBatchArr) {
	try {
		// #remove file all in folder
		let documentId = "";
		await Promise.all(
			dataBatchArr.map(async (value, i) => {
				const nameCopy = "copyTitle" + i;
				const copyDocId = await copyGoogleDoc(docId, nameCopy);
				const dataToReplace = await createArrayObjectToBatchUpdateDemo1(value);

				await writeDocBatchUpdate(copyDocId, dataToReplace);
				// await exportDocAsWord(copyDocId, nameCopy);
				// await exportGoogleDocAsPDF(copyDocId, nameCopy);
				// await deleteGoogleDoc(copyDocId);
				documentId = copyDocId;
				return copyDocId;
			})
		);
		return documentId;
	} catch (error) {
		console.error("Error copying document:", error);
		throw error;
	}
}

// === For Convert to PDF ===
async function exportDocAsWordDemo1(docId, setFileName) {
	const auth = await authorize();
	const drive = google.drive({ version: "v3", auth });
	const docs = google.docs({ version: "v1", auth });

	const PATH_FOLDER_KEEP_DOC_DEMO1 = `./download/doc_demo1/`;

	try {
		const res = await docs.documents.get({
			documentId: docId,
		});

		const fileId = res.data.documentId;
		const exportOptions = {
			mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		};

		const exportRes = await drive.files.export(
			{
				fileId: fileId,
				mimeType: exportOptions.mimeType,
			},
			{ responseType: "stream" }
		);

		const outputPath = path.join(process.cwd(), `${PATH_FOLDER_KEEP_DOC_DEMO1}${setFileName}.docx`);
		const fileStream = createWriteStream(outputPath);
		exportRes.data.pipe(fileStream);

		return new Promise((resolve, reject) => {
			fileStream.on("finish", resolve);
			fileStream.on("error", reject);
		});
	} catch (error) {
		console.error("Error exporting the document:", error);
	}
}

async function rebuildDocsDemo1ExportSaveToDrive(docId, dataBatchArr, setFileName, driveId) {
	try {
		let pdfPath;
		// #remove file all in folder
		const PATH_FOLDER_KEEP_DOC_DEMO1 = `./download/doc_demo1/`;
		await deleteFilesInFolderAsync(PATH_FOLDER_KEEP_DOC_DEMO1);

		const PATH_FOLDER_KEEP_PDF_DEMO1 = `./download/pdf_demo1/`;
		await deleteFilesInFolderAsync(PATH_FOLDER_KEEP_PDF_DEMO1);

		console.log("=== docId", docId);
		console.log("=== setFileName", setFileName);

		const fileId = await Promise.all(
			dataBatchArr.map(async (value, i) => {
				const copyDocId = await copyGoogleDoc(docId, setFileName);
				const dataToReplace = await createArrayObjectToBatchUpdateDemo1(value);

				console.log("=== writeDocBatchUpdate");
				await writeDocBatchUpdate(copyDocId, dataToReplace);

				console.log("=== exportDocAsWordDemo1");
				await exportDocAsWordDemo1(copyDocId, setFileName);

				/** === GOOGLE CONVERT PDF === */

				// console.log("=== exportGoogleDocAsPDF")
				// const pathSaveFilePDF = `./download/pdf_demo1/${setFileName}.pdf`
				// console.log("===  pathSaveFilePDF ", pathSaveFilePDF)

				// const fileId = await exportGoogleDocAsPDF(copyDocId, pathSaveFilePDF).then(async () => {
				// 	console.log("=== uploadFIleToDriveDemo1")
				// 	const fileId = await uploadFileToDriveDemo1(pathSaveFilePDF, setFileName, driveId);
				// 	await deleteGoogleDoc(copyDocId);
				// 	return fileId
				// })
				// return fileId

				/**=== NODE CONVERT PDF ===*/

				pdfPath = await convertDocxToPDF(setFileName);
				await uploadFileToDriveDemo1(pdfPath, setFileName, driveId);
				await deleteGoogleDoc(copyDocId);
			})
		);

		/** === GOOGLE CONVERT PDF === */

		// console.log("%c  === ","color:cyan","  fileId", fileId)
		// if (!fileId) throwError(STATUS_CODES.BAD_REQUEST, `can't get pdf path for upload file`);
		// return fileId;

		/**=== NODE CONVERT PDF ===*/

		console.log("%c  === ", "color:cyan", "  pdfPath", pdfPath);
		if (!pdfPath) throwError(STATUS_CODES.BAD_REQUEST, `can't get pdf path for upload file`);
		return pdfPath;
		
	} catch (error) {
		console.error("Error copying document:", error);
		throw error;
	}
}

async function uploadFileToDriveDemo1(fileUploadName, setFileName, driveId) {
	console.log("%c  === ", "color:cyan", "  fileUploadName", fileUploadName);

	try {
		if (!fileUploadName) throw Error("FilePath for upload invalid");
		const filePath = fileUploadName; // Replace with the path of the file to upload
		// const folderId = "1iQJf_x8zuHKxqyB1M-wvyfJJD3Pu7rMZ";
		const folderId = driveId;

		console.log("uploadFileToDriveDemo1  folderId", folderId);
		const fileName = setFileName; // Specify the name of the file in Google Drive
		const fileId = await uploadFileToDrive(filePath, folderId, fileName);
		console.log("File uploaded to Google Drive. File ID:", fileId);
		return fileId;
	} catch (error) {
		console.error(error);
		throw error;
	}
}

export {
	printDocTitle,
	copyGoogleDoc,
	rebuildDocs,
	copyGoogleSheet,
	createNewSpreadsheetWithData,
	downloadGoogleSheetsXlsx,
	downloadGoogleSheetsXlsxPipe,
	uploadFileToDriveDemo1,
	rebuildDocsDemo1,
	downloadGoogleDocsDocxPipe,
	rebuildDocsDemo1ExportSaveToDrive,
};

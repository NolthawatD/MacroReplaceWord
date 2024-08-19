import { google } from "googleapis";

// #Setting configuration
const getGoogleAuthClient = async () => {
	const auth = new google.auth.GoogleAuth({
		keyFile: "./src/config/credentials-sheet.json",
		scopes: "https://www.googleapis.com/auth/spreadsheets",
	});

	return await auth.getClient();
};

const getGoogleSheetsInstance = (authClient) => {
	return google.sheets({ version: "v4", auth: authClient });
};
// #End

// #Get MetaData Sheet
const useGoogleSheetsInstance = async () => {
	const authClient = await getGoogleAuthClient();
	const sheetsInstance = getGoogleSheetsInstance(authClient);
	return sheetsInstance;
};

const googleSheetsMetaData = async (spreadsheetId) => {
	try {
		// await createFileService()
		const googleSheets = await useGoogleSheetsInstance();
		const metaData = await googleSheets.spreadsheets.get({
			spreadsheetId,
		});
		return metaData;
	} catch (error) {
		// console.error("Error executing Google Sheets operations:", error);
		throw error;
	}
};

// #Check SheetExisting
const checkSheetExistence = async (spreadsheetId, sheetName) => {
	try {
		const spreadsheet = await googleSheetsMetaData(spreadsheetId);
		const sheets = spreadsheet.data.sheets;

		for (const sheet of sheets) {
			if (sheet.properties.title === sheetName) {
				return true;
			}
		}

		return false;
	} catch (error) {
		console.error("Error executing Google Sheets operations:", error);
		throw error;
	}
};

// #Get RowsFromSheet
const getRowsFromSheet = async (spreadsheetId, sheetName) => {
	try {
		const googleSheets = await useGoogleSheetsInstance();

		const response = await googleSheets.spreadsheets.values.get({
			spreadsheetId,
			range: sheetName,
		});

		const rows = response.data.values ? response.data.values : [];
		return rows;
	} catch (error) {
		// console.error("Error getting rows from sheet:", error);
		throw error;
	}
};

const getLengthRowsInSheet = async (spreadsheetId, sheetName) => {
	try {
		const rows = await getRowsFromSheet(spreadsheetId, sheetName);
		return rows.length;
	} catch (error) {
		// console.error("Error getting rows from sheet:", error);
		throw error;
	}
};

// Create a new spreadsheet
const createNewSpreadsheet = async (title) => {
	try {
		const googleSheets = await useGoogleSheetsInstance();
		const response = await googleSheets.spreadsheets.create({
			resource: {
				properties: {
					title: title,
				},
			},
		});

		// After creating the spreadsheet, make it public
		return response.data;
	} catch (error) {
		console.error("Error creating new spreadsheet:", error);
		throw error;
	}
};

export { googleSheetsMetaData, checkSheetExistence, getRowsFromSheet, getLengthRowsInSheet, createNewSpreadsheet };

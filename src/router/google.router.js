import express from "express";

import { sendResponse } from "../util/response.js";
import { STATUS_CODES } from "../util/status-code.js";
import { throwError } from "../util/throw-error.js";
import {
	createNewSpreadsheetWithData,
	downloadGoogleDocsDocxPipe,
	downloadGoogleSheetsXlsxPipe,
	printDocTitle,
	rebuildDocs,
	rebuildDocsDemo1,
	rebuildDocsDemo1ExportSaveToDrive,
} from "../service/google-api.js";
import { checkSheetExistence, getRowsFromSheet, googleSheetsMetaData } from "../service/google-sheet-api.js";

import path from "path";
import archiver from "archiver";
import { calculateRange, columnFormatData } from "../util/format-data.js";
import { convertStringToDate, formatDateThaiBuddhist } from "../util/format-date.js";
import { convertJSONFileToObject, convertObjectToJSONFile } from "../util/convert-json.js";
import { deleteFilesAsync, deleteFilesInFolderAsync } from "../util/fs-handler.js";

const PATH_FOLDER_KEEP_DOC = `./download/doc/`;
export const googleDocRouter = express.Router();

/**
 * @swagger
 * /isHasDoc:
 *   post:
 *     summary: Check if a Google Document exists
 *     description: Check whether a specified Google Document exists based on the provided documentId.
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               documentId:
 *                 type: string
 *                 description: The ID of the Google Document.
 *                 example: "1Ly7wCkaS-4dgZUlMkbokFjyzpcxYazdIS9BTOxWeYN0"
 *     responses:
 *       200:
 *         description: Successful response
 *         content:
 *           application/json:
 *             example: { result: true }
 *       400:
 *         description: Bad Request
 *         content:
 *           application/json:
 *             example: { result: false, message: "Required documentId" }
 *       404:
 *         description: Not Found
 *         content:
 *           application/json:
 *             example: { result: false, message: "Not found documentId: your-document-id" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { result: false, message: "An error occurred" }
 */
googleDocRouter.post("/isHasDoc", async (req, res) => {
	const { documentId } = req.body;
	try {
		if (!documentId) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required documentId");
		}

		const docTitle = await printDocTitle(documentId);
		if (!docTitle) return throwError(STATUS_CODES.NOT_FOUND, `Not found documentId: ${documentId}`);

		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", true);
	} catch (error) {
		if (error.throw) {
			return sendResponse(res, error.status, error.message, false);
		}
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /exportDoc:
 *   post:
 *     summary: Export a document with specified sheets
 *     description: |
 *      Export a Google Document based on the provided documentId and sheets array.
 *      Example sheets array:
 *         "documentId": "1Ly7wCkaS-4dgZUlMkbokFjyzpcxYazdIS9BTOxWeYN0",
 *         "sheetId": "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q",
 *         "sheetName": "investor",
 *         "sheetId": "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q",
 *         "sheetName": "Sheet2",
 *         "lengthRows": 5
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               documentId:
 *                 type: string
 *                 description: The ID of the Google Document.
 *                 example: "1Ly7wCkaS-4dgZUlMkbokFjyzpcxYazdIS9BTOxWeYN0"
 *               sheets:
 *                 type: array
 *                 description: An array of sheets to export.
 *                 items:
 *                   type: object
 *                   properties:
 *                     sheetId:
 *                       type: string
 *                       description: The ID of the sheet.
 *                       example: "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q"
 *                     sheetName:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "investor"
 *                     columnFormat:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "G"
 *                     columnsName:
 *                       type: array
 *                       description: An array of column names for the sheet.
 *                       items:
 *                         type: string
 *                         example: "columnsName"
 *               fromRow:
 *                 type: number
 *                 description: The from rows of want you use.
 *                 example: "1"
 *               toRow:
 *                 type: number
 *                 description: The row rows of want you use.
 *                 example: "5"
 *
 *     responses:
 *       200:
 *         description: Successful response
 *         content:
 *           application/json:
 *             example: { success: true }
 *       400:
 *         description: Bad Request
 *         content:
 *           application/json:
 *             example: { success: false, message: "Invalid or missing sheets array" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred" }
 */
googleDocRouter.post("/exportDoc", async (req, res) => {
	// case 1 row 1   fromRow == toRow
	// case 2 row all fromRow = 0 toRow = 0
	// case 3 custom    fromRow = 0 ถึง toRow = 0

	const { documentId, sheets, fromRow, toRow } = req.body;
	let _fromRow = Number(fromRow);
	let _toRow = Number(toRow);

	try {
		// #docs
		if (!documentId) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required documentId");
		}

		// #check documentId is existing
		const docTitle = await printDocTitle(documentId);
		if (!docTitle) return throwError(STATUS_CODES.NOT_FOUND, `Not found documentId: ${documentId}`);

		// #sheets
		if (!sheets || !Array.isArray(sheets)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Invalid or missing sheets array");
		}

		if (isNaN(_fromRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (isNaN(_toRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (_fromRow > _toRow) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow can't is more than toRow");
		}

		/**
		 * ถ้า request fromRow และ toRow ส่งมาเป็น ค่า 0 ให้ใช้แถวทั้งหมด
		 */

		console.log("%c  === ", "color:cyan", " start _fromRow === ", _fromRow);
		console.log("%c  === ", "color:cyan", " start _toRow === ", _toRow);

		// const uniqueColumns = [...new Set(sheets.flatMap((sheet) => sheet.columnsName))];
		// เอาหัวคอลัมน์ทั้งหมดมาสร้าง ไม่เลือกหัวคอลัมน์แล้้ว
		async function makeStructureArr(topicColumnArr, lengthRows) {
			console.log("%c  === ", "color:cyan", "  lengthRows", lengthRows);
			console.log("%c  === ", "color:cyan", "  topicColumnArr", topicColumnArr);

			if (topicColumnArr.length < 1) {
				return throwError(STATUS_CODES.BAD_REQUEST, "not found topic column");
			}

			const mergedArrStructToBatchUpdate = [];
			const emptyObject = {};
			topicColumnArr.forEach((column) => {
				emptyObject[column] = "";
			});

			for (let i = 0; i < lengthRows; i++) {
				mergedArrStructToBatchUpdate.push({ ...emptyObject });
			}

			console.log("%c  === ", "color:cyan", "  mergedArrStructToBatchUpdate", mergedArrStructToBatchUpdate);
			return mergedArrStructToBatchUpdate;
		}

		let mergedArrStructToBatchUpdate = [];

		// เอาแต่ละ sheet มาลูป หาคอลัมน์ในแต่ละ sheet *แต่มีชีทเดียวอยู่แล้วครับ
		for (const sheet of sheets) {
			if (!sheet.sheetId || !sheet.sheetName) {
				return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required sheetId and sheetName");
			}

			// เปลี่ยนเป็นเอาทุกหัวคอลัมน์
			// if (sheet.columnsName.length === 0) {
			// 	return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required column name must at least 1");
			// }

			// #check sheetId is existing
			const isSheetExist = await checkSheetExistence(sheet.sheetId, sheet.sheetName);
			if (!isSheetExist) {
				return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: '${sheet.sheetName}' not found`);
			}

			// #get rows data and check column name
			const rowsData = await getRowsFromSheet(sheet.sheetId, sheet.sheetName);
			if (rowsData.length === 1) {
				return throwError(STATUS_CODES.NOT_FOUND, `Rows length equals 1 can't do anymore`);
			}

			// #กรณีเลือก Select all จะเอาข้อมูลทั้งหมด
			if (_toRow === 0 && _fromRow === 0) {
				// ส่งหัวคอลัมน์ไป และจำนวนที่จะสร้าง
				mergedArrStructToBatchUpdate = await makeStructureArr(rowsData[0], rowsData.length - 1);
				_toRow = rowsData.length;
			}

			// #กรณีเลือก range custom หรือแถวเดียว
			if (_toRow > 0 && _fromRow > 0) {
				const lengthRows = _toRow - _fromRow + 1;
				if (rowsData.length < lengthRows) {
					return throwError(STATUS_CODES.NOT_FOUND, `Rows length have ${rowsData.length}, not enough`);
				}
				mergedArrStructToBatchUpdate = await makeStructureArr(rowsData[0], lengthRows);
			}

			console.log("%c  === ", "color:cyan", " End _fromRow === ", _fromRow);
			console.log("%c  === ", "color:cyan", " End _toRow === ", _toRow);

			if (_fromRow > 1) _fromRow -= 1;
			if (_fromRow === 0) _fromRow = 1;

			console.log("%c  === ", "color:cyan", " End 2 _fromRow === ", _fromRow);
			console.log("%c  === ", "color:cyan", " End 2 _toRow === ", _toRow);

			// push index 1 ใส่เป็นหัวให้ข้อมูลที่สไลด์ไปก่อนจะได้ข้อมูลที่น้อยลง ทำงานไวขึ้น
			let sliceRowsData = rowsData.slice(_fromRow, _toRow);
			console.log("%c  === ", "color:cyan", " sliceRowsData initial ", sliceRowsData);

			const $rowsDataWithTopic = [rowsData[0], ...sliceRowsData];
			console.log("%c  === ", "color:cyan", " $rowsDataWithTopic", $rowsDataWithTopic);

			if (sheet.columnFormat != "") {
				// #format data at column before sending and slice data already
				sliceRowsData = await columnFormatData(sheet.columnFormat, $rowsDataWithTopic);
				console.log("%c  === ", "color:cyan", "  sliceRowsData formatdata", sliceRowsData);
			}

			// #Push value to Sturct
			for (const colName of rowsData[0]) {
				// for (const colName of sheet.columnsName) {
				// #find columns where value is index ?
				const rowIndex = $rowsDataWithTopic[0].findIndex((element) => element === colName);
				console.log("%c  === ", "color:cyan", "  rowIndex", rowIndex);

				if (rowIndex === -1) {
					return throwError(STATUS_CODES.NOT_FOUND, `Not found column name: ${colName} at sheetId: ${sheet.sheetId}`);
				}

				await Promise.all(
					sliceRowsData.map((rowData, i) => {
						mergedArrStructToBatchUpdate[i][colName] = rowData[rowIndex];
					})
				);
			}
		}

		console.log("%c  === ", "color:cyan", "  mergedArrStructToBatchUpdate", mergedArrStructToBatchUpdate);

		// #send to batch update
		await rebuildDocs(documentId, mergedArrStructToBatchUpdate);
		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", true);
	} catch (error) {
		if (error.throw) {
			return sendResponse(res, error.status, error.message, false);
		}
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /exportSheet:
 *   post:
 *     summary: Export data from Google Sheets
 *     description: Export data from Google Sheets based on provided documentId, sheet information, and row range.
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               documentId:
 *                 type: string
 *                 description: The ID of the Google Document.
 *                 example: "1Ly7wCkaS-4dgZUlMkbokFjyzpcxYazdIS9BTOxWeYN0"
 *               sheets:
 *                 type: array
 *                 description: An array of sheets to export.
 *                 items:
 *                   type: object
 *                   properties:
 *                     sheetId:
 *                       type: string
 *                       description: The ID of the sheet.
 *                       example: "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q"
 *                     sheetName:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "Investor"
 *                     columnFormat:
 *                       type: string
 *                       description: The format of the columns.
 *                       example: "เลขบัตรประชาชน"
 *               fromRow:
 *                 type: integer
 *                 description: The starting row for exporting data. If set to 0, all rows will be used.
 *                 example: 1
 *               toRow:
 *                 type: integer
 *                 description: The ending row for exporting data. If set to 0, all rows will be used.
 *                 example: 5
 *     responses:
 *       '200':
 *         description: Data exported successfully
 *         content:
 *           application/json:
 *             example:
 *               success: true
 *       '400':
 *         description: Bad request
 *         content:
 *           application/json:
 *             example:
 *               success: false
 *               message: Invalid or missing sheets array
 *       '404':
 *         description: Resource not found
 *         content:
 *           application/json:
 *             example:
 *               success: false
 *               message: Sheet name not found
 *       '500':
 *         description: Internal server error
 *         content:
 *           application/json:
 *             example:
 *               success: false
 *               message: An error occurred
 */
googleDocRouter.post("/exportSheet", async (req, res) => {
	const { documentId, sheets, fromRow, toRow } = req.body;

	let _fromRow = Number(fromRow);
	let _toRow = Number(toRow);

	try {
		// #sheets
		if (!sheets || !Array.isArray(sheets)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Invalid or missing sheets array");
		}

		if (isNaN(_fromRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (isNaN(_toRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (_fromRow > _toRow) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow can't is more than toRow");
		}

		/**
		 * ถ้า request fromRow และ toRow ส่งมาเป็น ค่า 0 ให้ใช้คอลัมน์ทั้งหมด
		 */

		console.log("%c  === ", "color:cyan", "  _fromRow", _fromRow);
		console.log("%c  === ", "color:cyan", "  _toRow", _toRow);

		let newSpreadSheetId = "";

		// มีแค่ชีทเดียวที่ส่งมา
		for (const sheet of sheets) {
			if (!sheet.sheetId || !sheet.sheetName) {
				return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required sheetId and sheetName");
			}

			if (sheet.columnFormat === "") {
				return throwError(STATUS_CODES.BAD_REQUEST, "Required columnFormat");
			}

			// #check sheetId is existing
			const isSheetExist = await checkSheetExistence(sheet.sheetId, sheet.sheetName);
			if (!isSheetExist) {
				return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: '${sheet.sheetName}' not found`);
			}

			// #get rows data and check column name
			const rowsData = await getRowsFromSheet(sheet.sheetId, sheet.sheetName);
			if (rowsData.length <= 1) {
				return throwError(STATUS_CODES.NOT_FOUND, `Rows length equals 1 can't do anymore`);
			}

			// - - - Slice data - - -///

			if (_toRow === 0 && _fromRow === 0) {
				_toRow = rowsData.length;
			}

			// #กรณีเลือก range custom หรือแถวเดียว
			if (_toRow > 0 && _fromRow > 0) {
				const lengthRows = _toRow - _fromRow + 1;
				if (rowsData.length < lengthRows) {
					return throwError(STATUS_CODES.NOT_FOUND, `Rows length have ${rowsData.length}, not enough`);
				}
			}

			console.log("%c  === ", "color:cyan", " End _fromRow === ", _fromRow);
			console.log("%c  === ", "color:cyan", " End _toRow === ", _toRow);

			if (_fromRow > 1) _fromRow -= 1;
			if (_fromRow === 0) _fromRow = 1;

			console.log("%c  === ", "color:cyan", " End 2 _fromRow === ", _fromRow);
			console.log("%c  === ", "color:cyan", " End 2 _toRow === ", _toRow);

			// push index 1 ใส่เป็นหัวให้ข้อมูลที่สไลด์ไปก่อนจะได้ข้อมูลที่น้อยลง ทำงานไวขึ้น
			let sliceRowsData = rowsData.slice(_fromRow, _toRow);
			const $rowsDataWithTopic = [rowsData[0], ...sliceRowsData];

			// - - - Slice data - - -///

			// ส่งหัวคอลัมน์ไปด้วย ข้างในมีไสลด์ออกให้
			const newDataFormat = await columnFormatData(sheet.columnFormat, $rowsDataWithTopic);
			console.log("%c  === ", "color:cyan", "  newDataFormat", newDataFormat);

			const $newDataFormat = [rowsData[0], ...newDataFormat];

			const range = await calculateRange($newDataFormat);
			console.log("%c  === ", "color:cyan", "  range", range);

			//get name spreadsheet
			const spreadsheet = await googleSheetsMetaData(sheet.sheetId);
			console.log("%c  === ", "color:cyan", "  spreadsheet", spreadsheet.data.properties.title);
			const titleSpreadSheet = spreadsheet.data.properties.title + "-Format";

			const { spreadsheetId, writeData } = await createNewSpreadsheetWithData(titleSpreadSheet, $newDataFormat, `Sheet1!${range}`);
			console.log("%c  === ", "color:cyan", "  newSpreadSheet.status", writeData.status);
			if (writeData.status !== 200) {
				return throwError(STATUS_CODES.INTERNAL_SERVER_ERROR, `Failed to create new spreadsheet`);
			}

			console.log("%c  === ", "color:cyan", "  spreadsheetId", spreadsheetId);
			newSpreadSheetId = spreadsheetId;
		}

		// await downloadGoogleSheetsXlsxPipe(newSpreadSheetId, res, "setFileName");

		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", { spreadSheetId: newSpreadSheetId });
	} catch (error) {
		if (error.throw) {
			return sendResponse(res, error.status, error.message, false);
		}
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /download:
 *   get:
 *     summary: Download a folder as a ZIP file
 *     description: |
 *       Endpoint to download a folder as a ZIP file. Provide the path of the folder to be zipped.
 *     parameters:
 *       - in: query
 *         name: folderPath
 *         required: true
 *         schema:
 *           type: string
 *         description: The path of the folder to be zipped.
 *         example: "/path/to/folder"
 *     responses:
 *       200:
 *         description: Successful response with the ZIP file.
 *         content:
 *           application/zip:
 *             example: ZIP file binary data
 *       400:
 *         description: Bad Request
 *         content:
 *           application/json:
 *             example: { success: false, message: "Invalid folder path" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred during ZIP creation" }
 */
googleDocRouter.get("/download", async (req, res) => {
	const currentModuleUrl = new URL(import.meta.url);
	const currentModuleDir = path.dirname(currentModuleUrl.pathname);
	const twoStepsBack = path.join(currentModuleDir, "../..");
	const folderToZip = path.join(twoStepsBack, PATH_FOLDER_KEEP_DOC);

	// Create a ZIP archive
	const zip = archiver("zip");

	// Set the response headers for a zip file download
	res.attachment("data-folder.zip");

	// Pipe the zip archive to the response
	zip.pipe(res);

	// Add all files in the folder to the zip archive
	zip.directory(folderToZip, false);

	console.log("deleteFilesInFolderAsync successfully 1");

	// Event listener for when the zip process finishes
	zip.on("close", async function () {
		//   await deleteFilesInFolderAsync(folderToZip);
		console.log("deleteFilesInFolderAsync successfully 2");
		// Respond to the client with the ZIP file

		if (res.writableEnded) {
			console.log("Client download likely finished.");
		} else {
			console.log("Client download may not be finished.");
		}
	});

	// Finalize the zip archive
	zip.finalize();
});

/**
 * @swagger
 * /downloadSheetPipe:
 *   get:
 *     summary: Download Google Sheets as an XLSX file
 *     description: |
 *       Endpoint to download Google Sheets as an XLSX file. Provide the spreadsheet ID and set the file name.
 *     parameters:
 *       - in: query
 *         name: spreadSheetId
 *         required: true
 *         schema:
 *           type: string
 *         description: The ID of the Google Sheets document.
 *         example: "yourSpreadsheetId"
 *       - in: query
 *         name: setFileName
 *         required: true
 *         schema:
 *           type: string
 *         description: The name to be set for the downloaded file.
 *         example: "yourFileName.xlsx"
 *     responses:
 *       200:
 *         description: Successful response with the XLSX file.
 *         content:
 *           application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *             example: XLSX file binary data
 *       404:
 *         description: Not Found
 *         content:
 *           application/json:
 *             example: { success: false, message: "Resource not found" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred during file download" }
 */
googleDocRouter.get("/downloadSheetPipe", async (req, res) => {
	const { spreadSheetId, setFileName } = req.query;
	try {
		await downloadGoogleSheetsXlsxPipe(spreadSheetId, res, setFileName);
	} catch (error) {
		console.error(error);
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 *
 * 1. copy template by docID
 * 2. make new about format data special
 * 3. pick some field to convert to new and add to new special format value
 *
 *
 * 2.1 if don't send about topic columns
 * 2.2 setting all about topic sheet1 and sheet2 index0
 * 2.3 check is duplicate and send error if found
 * 2.4 build mergeArray for prev rebuild
 *
 */

/**
 * @swagger
 * /exportDocDemo1:
 *   post:
 *     summary: Export a document with specified sheets
 *     description: |
 *      Export a Google Document based on the provided documentId and sheets array.
 *      Example sheets array:
 *         "documentId": "1Ly7wCkaS-4dgZUlMkbokFjyzpcxYazdIS9BTOxWeYN0",
 *         "sheetId": "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q",
 *         "sheetName": "investor",
 *         "sheetId": "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q",
 *         "sheetName": "Sheet2",
 *         "lengthRows": 5
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               documentId:
 *                 type: string
 *                 description: The ID of the Google Document.
 *                 example: "1IbS-iVIpNbKRke6_fxb5lMwodMj8muj1123D8e3q6Zk"
 *               sheets:
 *                 type: array
 *                 description: An array of sheets to export.
 *                 items:
 *                   type: object
 *                   properties:
 *                     sheetId:
 *                       type: string
 *                       description: The ID of the sheet.
 *                       example: "16cZb_KaasnIXJlZNTaijGDrREpPMRGQGlMiC84aJZc8"
 *                     sheetName:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "Investor"
 *                     columnFormat:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "G"
 *                     columnsName:
 *                       type: array
 *                       description: An array of column names for the sheet.
 *                       items:
 *                         type: string
 *                         example: "columnsName"
 *               fromRow:
 *                 type: number
 *                 description: The from rows of want you use.
 *                 example: "1"
 *               toRow:
 *                 type: number
 *                 description: The row rows of want you use.
 *                 example: "5"
 *
 *     responses:
 *       200:
 *         description: Successful response
 *         content:
 *           application/json:
 *             example: { success: true }
 *       400:
 *         description: Bad Request
 *         content:
 *           application/json:
 *             example: { success: false, message: "Invalid or missing sheets array" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred" }
 */
googleDocRouter.post("/exportDocDemo1", async (req, res) => {
	const { documentId, sheets, fromRow, toRow, contractType } = req.body;
	// console.log("%c  === ","color:cyan","  contractType", contractType)
	// console.log("%c  === ", "color:cyan", "  lengthRows", lengthRows);
	// console.log("%c  === ", "color:cyan", "  lengthRows", typeof lengthRows);
	let _fromRow = Number(fromRow);
	let _toRow = Number(toRow);

	try {
		// #docs
		if (!documentId) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required documentId");
		}

		if (!["C2", "C3", "C4"].includes(contractType)) {
			return throwError(STATUS_CODES.BAD_REQUEST, `${contractType} is invalid contractType`);
		}

		// #check documentId is existing
		const docTitle = await printDocTitle(documentId);
		if (!docTitle) return throwError(STATUS_CODES.NOT_FOUND, `Not found documentId: ${documentId}`);

		// #sheets
		if (!sheets || !Array.isArray(sheets)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Invalid or missing sheets array");
		}

		if (isNaN(_fromRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (isNaN(_toRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (_fromRow > _toRow) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow can't is more than toRow");
		}

		if (_fromRow < 2 || _toRow < 2) {
			return throwError(STATUS_CODES.BAD_REQUEST, `Row have to more than 1`);
		}

		if (_fromRow < 2 || _toRow < 2) {
			return throwError(STATUS_CODES.BAD_REQUEST, `fromRow is ${_fromRow} or toRow is ${_toRow} both have to more than 2`);
		}

		// case pick a row
		if (_fromRow !== _toRow) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Internal server error _fromRow not equal _toRow");
		}

		/**
		 * ถ้า request fromRow และ toRow ส่งมาเป็น ค่า 0 ให้ใช้คอลัมน์ทั้งหมด
		 */

		// console.log("%c  === ", "color:cyan", "  _fromRow", _fromRow);
		// console.log("%c  === ", "color:cyan", "  _toRow", _toRow);

		console.log("%c  === ", "color:cyan", "  sheets[0]?.columnsName[0]", sheets[0]?.columnsName[0]);
		let buildColumnsNameArr = [];

		// check if is column name is undefine will be get all column one and two build mergeArrayBuild
		if (!sheets[0]?.columnsName[0]) {
			for (const sheet of sheets) {
				if (!sheet.sheetId || !sheet.sheetName) {
					return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required sheetId and sheetName");
				}

				// #check sheetId is existing
				const isSheetExist = await checkSheetExistence(sheet.sheetId, sheet.sheetName);
				if (!isSheetExist) {
					return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: '${sheet.sheetName}' not found`);
				}
				const rowsData = await getRowsFromSheet(sheet.sheetId, sheet.sheetName);
				if (rowsData.length < 2) {
					return throwError(STATUS_CODES.NOT_FOUND, `Rows length equals 1 can't do anymore`);
				}
				console.log("%c  === ", "color:cyan", "  rowsData[0]", rowsData[0]);
				buildColumnsNameArr = [...buildColumnsNameArr, ...rowsData[0]];
			}
		}

		// #check duplucate buildColumnsNameArr
		console.log("%c  === ", "color:cyan", "  buildColumnsNameArr", buildColumnsNameArr);
		if (buildColumnsNameArr.length > 0) {
			const encounteredValues = {};
			buildColumnsNameArr.forEach((value) => {
				if (encounteredValues[value]) {
					console.error(`Duplicate columns name value found: ${value}`);
					return throwError(STATUS_CODES.NOT_FOUND, `Duplicate columns name value found: ${value}`);
				} else {
					encounteredValues[value] = true;
				}
			});
		}

		async function makeStructureArr(columnNameArr, lengthRows) {
			const mergedArrStructToBatchUpdate = [];
			const emptyObject = {};
			columnNameArr.forEach((column) => {
				emptyObject[column] = "";
			});

			for (let i = 0; i < lengthRows; i++) {
				mergedArrStructToBatchUpdate.push({ ...emptyObject });
			}
			console.log("%c  === ", "color:cyan", "  mergedArrStructToBatchUpdate", mergedArrStructToBatchUpdate);
			return mergedArrStructToBatchUpdate;
		}

		let mergedArrStructToBatchUpdate = [];
		if (buildColumnsNameArr.length > 0) {
			console.log("=== **ไม่มีการส่งคอลัทน์เนมเข้ามาอยู่แล้ว** ===");
			mergedArrStructToBatchUpdate = await makeStructureArr(buildColumnsNameArr, 1);
		}

		// const uniqueColumns = [...new Set(sheets.flatMap((sheet) => sheet.columnsName))];
		// if (uniqueColumns.length > 0 && buildColumnsNameArr.length > 0) {
		// 	console.log("=== มีการส่งคอลัทน์เนมเข้ามา ===");
		// 	mergedArrStructToBatchUpdate = await makeStructureArr(buildColumnsNameArr, 1);
		// }

		// เอาแต่ละ sheet มาลูป หาคอลัมน์ในแต่ละ sheet
		let totalLoop = 1;
		for (const sheet of sheets) {
			if (!sheet.sheetId || !sheet.sheetName) {
				return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required sheetId and sheetName");
			}

			// #check sheetId is existing
			const isSheetExist = await checkSheetExistence(sheet.sheetId, sheet.sheetName);
			if (!isSheetExist) {
				return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: '${sheet.sheetName}' not found`);
			}

			// #get rows data and check column name
			const rowsData = await getRowsFromSheet(sheet.sheetId, sheet.sheetName);
			if (rowsData.length === 1) {
				return throwError(STATUS_CODES.NOT_FOUND, `Rows length equals 1 can't do anymore`);
			}

			if (_toRow > 0) {
				const lengthRows = _toRow - _fromRow + 1;
				if (rowsData.length < lengthRows) {
					return throwError(STATUS_CODES.NOT_FOUND, `Rows length have ${rowsData.length}, not enough`);
				}
			}

			console.log("%c  === ", "color:cyan", "  _fromRow", _fromRow - 1);
			console.log("%c  === ", "color:cyan", "  _toRow", _toRow);

			let sliceRowsData = [];
			let $rowsDataWithTopic = [];

			console.log("%c  === ", "color:cyan", "  totalLoop", totalLoop);

			if (totalLoop === 1) {
				// push index 1 ใส่เป็นหัวให้ข้อมูลที่สไลด์ไปก่อนจะได้ข้อมูลที่น้อยลง ทำงานไวขึ้น
				sliceRowsData = rowsData.slice(_fromRow - 1, _toRow);
				console.log("%c  === ", "color:cyan", " sliceRowsData initial ", sliceRowsData);

				$rowsDataWithTopic = [rowsData[0], ...sliceRowsData];
				console.log("%c  === ", "color:cyan", " $rowsDataWithTopic", $rowsDataWithTopic);
			}

			if (totalLoop === 2) {
				const smeMapping = mergedArrStructToBatchUpdate[0]["จับคู่กับนิติบุคคล (เลขแถว)"]; //// may get request from client
				console.log("%c  === ", "color:cyan", "  smeMapping", smeMapping);

				if (!smeMapping) {
					return throwError(STATUS_CODES.NOT_FOUND, `smeMapping is ${smeMapping} that mismatch or invalid data`);
				}

				if (Number(smeMapping) > rowsData.length) {
					return throwError(
						STATUS_CODES.NOT_FOUND,
						`smeMapping is ${Number(smeMapping)} more than length of rowsData (${rowsData.length})`
					);
				}

				sliceRowsData = rowsData.slice(smeMapping - 1, smeMapping);
				console.log("%c  === ", "color:cyan", " sliceRowsData initial ", sliceRowsData);

				$rowsDataWithTopic = [rowsData[0], ...sliceRowsData];
				console.log("%c  === ", "color:cyan", " $rowsDataWithTopic", $rowsDataWithTopic);
			}

			// #keep value to Sturct
			// #loop for get index
			if (buildColumnsNameArr.length > 0) {
				for (const colName of rowsData[0]) {
					const rowIndex = $rowsDataWithTopic[0].findIndex((element) => element === colName);
					// console.log("%c  === ", "color:cyan", "  rowIndex", rowIndex);

					if (rowIndex === -1) {
						console.log(`Not found column name: ${colName} at sheetId: ${sheet.sheetId}`);
						return throwError(STATUS_CODES.NOT_FOUND, `Not found column name: ${colName} at sheetId: ${sheet.sheetId}`);
					}

					// #push data to merge
					await Promise.all(
						sliceRowsData.map((rowData, i) => {
							mergedArrStructToBatchUpdate[i][colName] = rowData[rowIndex];
						})
					);
				}
			}

			totalLoop++;
		}

		console.log("%c  === ", "color:cyan", " ==> prepare mergedArrStructToBatchUpdate before", mergedArrStructToBatchUpdate);

		/////// -> #convert data to new again it be easy to send to rebuildDocs
		const documentType = contractType; // C2, C3, C4
		if (documentType === "C2") {
			// add key to
			const startContractDate = mergedArrStructToBatchUpdate[0]["วันที่เริ่มปันผล"];
			const numberOfMonths = Number(mergedArrStructToBatchUpdate[0]["จำนวนเดือน ตามสัญญา (ใส่เป็นเลขเท่านั้น)"]);
			const numberDatePaid = mergedArrStructToBatchUpdate[0]["วันที่จ่ายปันผล"];
			const interestMonthPaid = mergedArrStructToBatchUpdate[0]["ดอกเบี้ยปันผลต่อเดือน"];

			console.log("%c  === ", "color:cyan", "  startContractDate :", startContractDate);
			console.log("%c  === ", "color:cyan", "  numberOfMonths : ", numberOfMonths);
			console.log("%c  === ", "color:cyan", "  numberDatePaid : ", numberDatePaid);
			console.log("%c  === ", "color:cyan", "  interestMonthPaid : ", interestMonthPaid);

			const proceedPatternForm = [];
			const startDate = convertStringToDate(startContractDate);

			// Loop through the months
			for (let i = 0; i < numberOfMonths; i++) {
				const currentMonth = (startDate.getMonth() + i) % 12; // Corrected line
				let currentYear = startDate.getFullYear() + Math.floor((startDate.getMonth() + i) / 12);
				const daysInCurrentMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
				let _numberDatePaid = numberDatePaid;

				if (currentMonth === 1 && numberDatePaid > 28) {
					_numberDatePaid = daysInCurrentMonth;
				}

				// Adjust the year if the current month exceeds December
				if (currentMonth >= 12) {
					currentYear += 1;
				}

				const currentDate = new Date(currentYear, currentMonth, Math.min(startDate.getDate(), _numberDatePaid));

				// numberDatePaid if date is 30 on Feb is make change to last day of this month
				const patternForm = `\tข้อ 6.${i + 1}) ชำระดอกเบี้ยเงินกู้ รอบเดือน ${formatDateThaiBuddhist(
					currentDate
				)} คือ วันที่ ${_numberDatePaid} จำนวน ${interestMonthPaid} บาท\n`;
				proceedPatternForm.push(patternForm);
			}

			mergedArrStructToBatchUpdate[0]["ประมวลผลตามรอบเดือนปันผล"] = proceedPatternForm;
		}

		console.log("%c  === ", "color:cyan", " ==> prepare mergedArrStructToBatchUpdate after", mergedArrStructToBatchUpdate);

		// #send to batch update

		function replaceWhiteSpaceWithDash(inputString) {
			console.log("%c  === ", "color:cyan", "  inputString", inputString);
			return inputString.replace(/\s+/g, "");
		}

		// Setting file name
		const businessName = mergedArrStructToBatchUpdate[0]["ชื่อสกุล ผปก"];
		if (!businessName) {
			return throwError(STATUS_CODES.NOT_FOUND, `Not found ชื่อ สกุล ผปก`);
		}
		const investorName = mergedArrStructToBatchUpdate[0]["ชื่อสกุล นลท"];
		if (!businessName) {
			return throwError(STATUS_CODES.NOT_FOUND, `Not found ชื่อสกุล นลท`);
		}
		const _businessName = replaceWhiteSpaceWithDash(businessName);
		const _investorName = replaceWhiteSpaceWithDash(investorName);

		const fileName = `${documentType}_${_businessName}_${_investorName}`;

		await convertObjectToJSONFile(mergedArrStructToBatchUpdate, fileName);
		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", { fileName, businessName, investorName });
	} catch (error) {
		if (error.throw) {
			return sendResponse(res, error.status, error.message, false);
		}
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /buildDocumentDemo1:
 *   get:
 *     summary: Download Google Docs as a DOCX file
 *     description: |
 *       Endpoint to download Google Docs as a DOCX file. Provide the document ID and set the file name.
 *     parameters:
 *       - in: query
 *         name: documentId
 *         required: true
 *         schema:
 *           type: string
 *         description: The ID of the Google Docs document.
 *         example: "yourDocumentId"
 *       - in: query
 *         name: setFileName
 *         required: true
 *         schema:
 *           type: string
 *         description: The name to be set for the downloaded file.
 *         example: "yourFileName.docx"
 *     responses:
 *       200:
 *         description: Successful response with the DOCX file.
 *         content:
 *           application/vnd.openxmlformats-officedocument.wordprocessingml.document:
 *             example: DOCX file binary data
 *       404:
 *         description: Not Found
 *         content:
 *           application/json:
 *             example: { success: false, message: "Resource not found" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred during file download" }
 */
googleDocRouter.get("/buildDocumentDemo1", async (req, res) => {
	const { documentId, setFileName } = req.query;
	console.log(" ==== Buliding Doc ===");
	console.log("%c  === ", "color:cyan", "  setFileName", setFileName);
	console.log("%c  === ", "color:cyan", "  documentId", documentId);

	try {
		if (!documentId) return sendResponse(res, STATUS_CODES.BAD_REQUEST || 404, `documentId = ${documentId} invalid`, false);
		if (!setFileName) return sendResponse(res, STATUS_CODES.BAD_REQUEST || 404, `setFileName = ${setFileName} invalid`, false);

		// convert to object to send rebuild
		const objectRebuild = JSON.parse(await convertJSONFileToObject(setFileName));
		console.log("%c  === ", "color:cyan", "  objectRebuild", typeof objectRebuild);

		const buildDocumentId = await rebuildDocsDemo1(documentId, objectRebuild);
		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", { documentId: buildDocumentId, setFileName: setFileName });
	} catch (error) {
		console.error(error);
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /downloadDocPipe:
 *   get:
 *     summary: Download Google Docs as a DOCX file
 *     description: |
 *       Endpoint to download Google Docs as a DOCX file. Provide the document ID and set the file name.
 *     parameters:
 *       - in: query
 *         name: documentId
 *         required: true
 *         schema:
 *           type: string
 *         description: The ID of the Google Docs document.
 *         example: "yourDocumentId"
 *       - in: query
 *         name: setFileName
 *         required: true
 *         schema:
 *           type: string
 *         description: The name to be set for the downloaded file.
 *         example: "yourFileName.docx"
 *     responses:
 *       200:
 *         description: Successful response with the DOCX file.
 *         content:
 *           application/vnd.openxmlformats-officedocument.wordprocessingml.document:
 *             example: DOCX file binary data
 *       404:
 *         description: Not Found
 *         content:
 *           application/json:
 *             example: { success: false, message: "Resource not found" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred during file download" }
 */
googleDocRouter.get("/downloadDocPipe", async (req, res) => {
	/// documentId for rebuild
	/// setFileName name form data exportDocDemo1
	const { documentId, setFileName } = req.query;
	console.log(" ==== Download Doc ===");
	console.log("%c  === ", "color:cyan", "  setFileName", setFileName);
	console.log("%c  === ", "color:cyan", "  documentId", documentId);

	try {
		if (!documentId) return sendResponse(res, STATUS_CODES.BAD_REQUEST || 404, `documentId = ${documentId} invalid`, false);
		if (!setFileName) return sendResponse(res, STATUS_CODES.BAD_REQUEST || 404, `setFileName = ${setFileName} invalid`, false);

		await downloadGoogleDocsDocxPipe(documentId, res, "fileNameKeepToDrive");

		console.log("=== Delete File Temp ===");
		const currentModuleUrl = new URL(import.meta.url);
		const currentModuleDir = path.dirname(currentModuleUrl.pathname);
		const projectDir = path.join(currentModuleDir, "../..");
		await deleteFilesAsync(projectDir + `/temp/${setFileName}.json`);
	} catch (error) {
		console.error(error);
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /saveToDriveDemo1:
 *   post:
 *     summary: Export a document with specified sheets
 *     description: |
 *      Export a Google Document based on the provided documentId and sheets array.
 *      Example sheets array:
 *         "documentId": "1Ly7wCkaS-4dgZUlMkbokFjyzpcxYazdIS9BTOxWeYN0",
 *         "sheetId": "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q",
 *         "sheetName": "investor",
 *         "sheetId": "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q",
 *         "sheetName": "Sheet2",
 *         "lengthRows": 5
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               documentId:
 *                 type: string
 *                 description: The ID of the Google Document.
 *                 example: "1IbS-iVIpNbKRke6_fxb5lMwodMj8muj1123D8e3q6Zk"
 *               sheets:
 *                 type: array
 *                 description: An array of sheets to export.
 *                 items:
 *                   type: object
 *                   properties:
 *                     sheetId:
 *                       type: string
 *                       description: The ID of the sheet.
 *                       example: "16cZb_KaasnIXJlZNTaijGDrREpPMRGQGlMiC84aJZc8"
 *                     sheetName:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "Investor"
 *                     columnFormat:
 *                       type: string
 *                       description: The name of the sheet.
 *                       example: "G"
 *                     columnsName:
 *                       type: array
 *                       description: An array of column names for the sheet.
 *                       items:
 *                         type: string
 *                         example: "columnsName"
 *               fromRow:
 *                 type: number
 *                 description: The from rows of want you use.
 *                 example: "1"
 *               toRow:
 *                 type: number
 *                 description: The row rows of want you use.
 *                 example: "5"
 *
 *     responses:
 *       200:
 *         description: Successful response
 *         content:
 *           application/json:
 *             example: { success: true }
 *       400:
 *         description: Bad Request
 *         content:
 *           application/json:
 *             example: { success: false, message: "Invalid or missing sheets array" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { success: false, message: "An error occurred" }
 */
googleDocRouter.post("/saveToDriveDemo1", async (req, res) => {
	const { documentId, sheets, fromRow, toRow, contractType } = req.body;

	let _fromRow = Number(fromRow);
	let _toRow = Number(toRow);

	try {
		// #docs
		if (!documentId) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required documentId");
		}

		if (!["C2", "C3", "C4"].includes(contractType)) {
			return throwError(STATUS_CODES.BAD_REQUEST, `${contractType} is invalid contractType`);
		}

		// #check documentId is existing
		const docTitle = await printDocTitle(documentId);
		if (!docTitle) return throwError(STATUS_CODES.NOT_FOUND, `Not found documentId: ${documentId}`);

		// #sheets
		if (!sheets || !Array.isArray(sheets)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Invalid or missing sheets array");
		}

		if (isNaN(_fromRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (isNaN(_toRow)) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow is not a number");
		}

		if (_fromRow > _toRow) {
			return throwError(STATUS_CODES.BAD_REQUEST, "fromRow can't is more than toRow");
		}

		if (_fromRow < 2 || _toRow < 2) {
			return throwError(STATUS_CODES.BAD_REQUEST, `Row have to more than 1`);
		}

		if (_fromRow < 2 || _toRow < 2) {
			return throwError(STATUS_CODES.BAD_REQUEST, `fromRow is ${_fromRow} or toRow is ${_toRow} both have to more than 2`);
		}

		// case pick a row
		if (_fromRow !== _toRow) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Internal server error _fromRow not equal _toRow");
		}

		/**
		 * ถ้า request fromRow และ toRow ส่งมาเป็น ค่า 0 ให้ใช้คอลัมน์ทั้งหมด
		 */

		// console.log("%c  === ", "color:cyan", "  _fromRow", _fromRow);
		// console.log("%c  === ", "color:cyan", "  _toRow", _toRow);

		console.log("%c  === ", "color:cyan", "  sheets[0]?.columnsName[0]", sheets[0]?.columnsName[0]);
		let buildColumnsNameArr = [];

		// check if is column name is undefine will be get all column one and two build mergeArrayBuild
		if (!sheets[0]?.columnsName[0]) {
			for (const sheet of sheets) {
				if (!sheet.sheetId || !sheet.sheetName) {
					return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required sheetId and sheetName");
				}

				// #check sheetId is existing
				const isSheetExist = await checkSheetExistence(sheet.sheetId, sheet.sheetName);
				if (!isSheetExist) {
					return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: '${sheet.sheetName}' not found`);
				}
				const rowsData = await getRowsFromSheet(sheet.sheetId, sheet.sheetName);
				if (rowsData.length < 2) {
					return throwError(STATUS_CODES.NOT_FOUND, `Rows length equals 1 can't do anymore`);
				}
				console.log("%c  === ", "color:cyan", "  rowsData[0]", rowsData[0]);
				buildColumnsNameArr = [...buildColumnsNameArr, ...rowsData[0]];
			}
		}

		// #check duplucate buildColumnsNameArr
		console.log("%c  === ", "color:cyan", "  buildColumnsNameArr", buildColumnsNameArr);
		if (buildColumnsNameArr.length > 0) {
			const encounteredValues = {};
			buildColumnsNameArr.forEach((value) => {
				if (encounteredValues[value]) {
					console.error(`Duplicate columns name value found: ${value}`);
					return throwError(STATUS_CODES.NOT_FOUND, `Duplicate columns name value found: ${value}`);
				} else {
					encounteredValues[value] = true;
				}
			});
		}

		async function makeStructureArr(columnNameArr, lengthRows) {
			const mergedArrStructToBatchUpdate = [];
			const emptyObject = {};
			columnNameArr.forEach((column) => {
				emptyObject[column] = "";
			});

			for (let i = 0; i < lengthRows; i++) {
				mergedArrStructToBatchUpdate.push({ ...emptyObject });
			}
			console.log("%c  === ", "color:cyan", "  mergedArrStructToBatchUpdate", mergedArrStructToBatchUpdate);
			return mergedArrStructToBatchUpdate;
		}

		let mergedArrStructToBatchUpdate = [];
		if (buildColumnsNameArr.length > 0) {
			console.log("=== ไม่มีการส่งคอลัทน์เนมเข้ามา ===");
			mergedArrStructToBatchUpdate = await makeStructureArr(buildColumnsNameArr, 1);
		}

		const uniqueColumns = [...new Set(sheets.flatMap((sheet) => sheet.columnsName))];
		if (uniqueColumns.length > 0 && buildColumnsNameArr.length > 0) {
			console.log("=== มีการส่งคอลัทน์เนมเข้ามา ===");
			mergedArrStructToBatchUpdate = await makeStructureArr(buildColumnsNameArr, 1);
		}

		// เอาแต่ละ sheet มาลูป หาคอลัมน์ในแต่ละ sheet
		let totalLoop = 1;
		for (const sheet of sheets) {
			if (!sheet.sheetId || !sheet.sheetName) {
				return throwError(STATUS_CODES.BAD_REQUEST, "Invalid sheet object in the array, Required sheetId and sheetName");
			}

			// #check sheetId is existing
			const isSheetExist = await checkSheetExistence(sheet.sheetId, sheet.sheetName);
			if (!isSheetExist) {
				return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: '${sheet.sheetName}' not found`);
			}

			// #get rows data and check column name
			const rowsData = await getRowsFromSheet(sheet.sheetId, sheet.sheetName);
			if (rowsData.length === 1) {
				return throwError(STATUS_CODES.NOT_FOUND, `Rows length equals 1 can't do anymore`);
			}

			if (_toRow > 0) {
				const lengthRows = _toRow - _fromRow + 1;
				if (rowsData.length < lengthRows) {
					return throwError(STATUS_CODES.NOT_FOUND, `Rows length have ${rowsData.length}, not enough`);
				}
			}

			console.log("%c  === ", "color:cyan", "  _fromRow", _fromRow - 1);
			console.log("%c  === ", "color:cyan", "  _toRow", _toRow);

			let sliceRowsData = [];
			let $rowsDataWithTopic = [];

			console.log("%c  === ", "color:cyan", "  totalLoop", totalLoop);

			if (totalLoop === 1) {
				// push index 1 ใส่เป็นหัวให้ข้อมูลที่สไลด์ไปก่อนจะได้ข้อมูลที่น้อยลง
				sliceRowsData = rowsData.slice(_fromRow - 1, _toRow);
				console.log("%c  === ", "color:cyan", " sliceRowsData initial ", sliceRowsData);

				$rowsDataWithTopic = [rowsData[0], ...sliceRowsData];
				console.log("%c  === ", "color:cyan", " $rowsDataWithTopic", $rowsDataWithTopic);
			}

			if (totalLoop === 2) {
				const smeMapping = mergedArrStructToBatchUpdate[0]["จับคู่กับนิติบุคคล (เลขแถว)"]; //// may get request from client
				console.log("%c  === ", "color:cyan", "  smeMapping", smeMapping);

				if (!smeMapping) {
					return throwError(STATUS_CODES.NOT_FOUND, `smeMapping is ${smeMapping} that mismatch or invalid data`);
				}

				if (Number(smeMapping) > rowsData.length) {
					return throwError(
						STATUS_CODES.NOT_FOUND,
						`smeMapping is ${Number(smeMapping)} more than length of rowsData (${rowsData.length})`
					);
				}

				sliceRowsData = rowsData.slice(smeMapping - 1, smeMapping);
				console.log("%c  === ", "color:cyan", " sliceRowsData initial ", sliceRowsData);

				$rowsDataWithTopic = [rowsData[0], ...sliceRowsData];
				console.log("%c  === ", "color:cyan", " $rowsDataWithTopic", $rowsDataWithTopic);
			}

			// #keep value to Sturct
			// #loop for get index
			if (buildColumnsNameArr.length > 0) {
				for (const colName of rowsData[0]) {
					const rowIndex = $rowsDataWithTopic[0].findIndex((element) => element === colName);
					// console.log("%c  === ", "color:cyan", "  rowIndex", rowIndex);

					if (rowIndex === -1) {
						console.log(`Not found column name: ${colName} at sheetId: ${sheet.sheetId}`);
						return throwError(STATUS_CODES.NOT_FOUND, `Not found column name: ${colName} at sheetId: ${sheet.sheetId}`);
					}

					// #push data to merge
					await Promise.all(
						sliceRowsData.map((rowData, i) => {
							mergedArrStructToBatchUpdate[i][colName] = rowData[rowIndex];
						})
					);
				}
			}

			totalLoop++;
		}

		console.log("%c  === ", "color:cyan", " ==> prepare mergedArrStructToBatchUpdate before", mergedArrStructToBatchUpdate);

		/////// -> #convert data to new again it be easy to send to rebuildDocs
		const documentType = contractType; // C2, C3, C4
		if (documentType === "C2") {
			// add key to
			const startContractDate = mergedArrStructToBatchUpdate[0]["วันที่เริ่มปันผล"];
			const numberOfMonths = Number(mergedArrStructToBatchUpdate[0]["จำนวนเดือน ตามสัญญา (ใส่เป็นเลขเท่านั้น)"]);
			const numberDatePaid = mergedArrStructToBatchUpdate[0]["วันที่จ่ายปันผล"];
			const interestMonthPaid = mergedArrStructToBatchUpdate[0]["ดอกเบี้ยปันผลต่อเดือน"];

			console.log("%c  === ", "color:cyan", "  startContractDate :", startContractDate);
			console.log("%c  === ", "color:cyan", "  numberOfMonths : ", numberOfMonths);
			console.log("%c  === ", "color:cyan", "  numberDatePaid : ", numberDatePaid);
			console.log("%c  === ", "color:cyan", "  interestMonthPaid : ", interestMonthPaid);

			const proceedPatternForm = [];
			const startDate = convertStringToDate(startContractDate);

			// Loop through the months
			for (let i = 0; i < numberOfMonths; i++) {
				const currentMonth = (startDate.getMonth() + i) % 12; // Corrected line
				let currentYear = startDate.getFullYear() + Math.floor((startDate.getMonth() + i) / 12);
				const daysInCurrentMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
				let _numberDatePaid = numberDatePaid;

				if (currentMonth === 1 && numberDatePaid > 28) {
					_numberDatePaid = daysInCurrentMonth;
				}

				// Adjust the year if the current month exceeds December
				if (currentMonth >= 12) {
					currentYear += 1;
				}

				const currentDate = new Date(currentYear, currentMonth, Math.min(startDate.getDate(), _numberDatePaid));

				// numberDatePaid if date is 30 on Feb is make change to last day of this month
				const patternForm = `\tข้อ 6.${i + 1}) ชำระดอกเบี้ยเงินกู้ รอบเดือน ${formatDateThaiBuddhist(
					currentDate
				)} คือ วันที่ ${_numberDatePaid} จำนวน ${interestMonthPaid} บาท\n`;
				proceedPatternForm.push(patternForm);
			}

			mergedArrStructToBatchUpdate[0]["ประมวลผลตามรอบเดือนปันผล"] = proceedPatternForm;
		}
		console.log("%c  === ", "color:cyan", " ==> prepare mergedArrStructToBatchUpdate after", mergedArrStructToBatchUpdate);

		function replaceWhiteSpaceWithDash(inputString) {
			return inputString.replace(/\s+/g, "");
		}

		// Setting file name
		// Setting file name
		const businessName = mergedArrStructToBatchUpdate[0]["ชื่อสกุล ผปก"];
		const investorName = mergedArrStructToBatchUpdate[0]["ชื่อสกุล นลท"];
		const _bussinessName = replaceWhiteSpaceWithDash(mergedArrStructToBatchUpdate[0]["ชื่อสกุล ผปก"]);
		const _investorName = replaceWhiteSpaceWithDash(mergedArrStructToBatchUpdate[0]["ชื่อสกุล นลท"]);
		const fileName = `${documentType}_${_bussinessName}_${_investorName}`;
		const pdfPath = await rebuildDocsDemo1ExportSaveToDrive(documentId, mergedArrStructToBatchUpdate, fileName);
		console.log("%c  === ", "color:cyan", "  pdfPath", pdfPath);
		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", { filePath: pdfPath });
	} catch (error) {
		if (error.throw) {
			return sendResponse(res, error.status, error.message, false);
		}
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

googleDocRouter.get("/saveToDriveDemo1V2", async (req, res) => {
	const { documentId, setFileName, driveId } = req.query;
	console.log(" ==== Buliding Doc ===");
	console.log("%c  === ", "color:cyan", "  setFileName", setFileName);
	console.log("%c  === ", "color:cyan", "  documentId", documentId);
	console.log("%c  === ","color:cyan","  driveId", driveId)

	try {
		if (!documentId) return sendResponse(res, STATUS_CODES.BAD_REQUEST || 404, `documentId = ${documentId} invalid`, false);
		if (!setFileName) return sendResponse(res, STATUS_CODES.BAD_REQUEST || 404, `setFileName = ${setFileName} invalid`, false);

		// convert to object to send rebuild
		const objectRebuild = JSON.parse(await convertJSONFileToObject(setFileName));
		console.log("%c  === ", "color:cyan", "  objectRebuild", typeof objectRebuild);

		const pdfPath = await rebuildDocsDemo1ExportSaveToDrive(documentId, objectRebuild, setFileName, driveId);
		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", { pdfPath });
	} catch (error) {
		console.error(error);
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

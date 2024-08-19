import express from "express";

import { checkSheetExistence, getRowsFromSheet, googleSheetsMetaData, getLengthRowsInSheet } from "../service/google-sheet-api.js";
import { sendResponse } from "../util/response.js";
import { STATUS_CODES } from "../util/status-code.js";
import { throwError } from "../util/throw-error.js";

// import * as AuthorService from "./author.service";

export const googleSheetRouter = express.Router();

googleSheetRouter.get("/googleMetaData", async (req, res) => {
	try {
		const { spreadSheetId } = req.query;
		if (!spreadSheetId) {
			return sendResponse(res, STATUS_CODES.BAD_REQUEST, "Required spreadSheetId", null);
		}
		const metaData = await googleSheetsMetaData(spreadSheetId);
		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", metaData.data);
	} catch (error) {
		return sendResponse(res, STATUS_CODES.INTERNAL_SERVER_ERROR, "Internal Server Error", null);
	}
});

/**
 * @swagger
 * /isHasSheet:
 *   post:
 *     summary: Check if a sheet exists in a Google Spreadsheet
 *     description: Check whether a specified sheet exists in a Google Spreadsheet.
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               spreadSheetId:
 *                 type: string
 *                 description: The ID of the Google Spreadsheet.
 *                 example: "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q"
 *               sheetName:
 *                 type: string
 *                 description: The name of the sheet to check for existence.
 *                 example: "investor"
 *               lengthRows:
 *                 type: number
 *                 description: The length rows of the sheet to check for enough.
 *                 example: "10"
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
 *             example: { status: "Bad Request", message: "Required spreadSheetId or sheetName", result: false  }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { status: "Internal Server Error", message: "An error occurred", result: false  }
 */
googleSheetRouter.post("/isHasSheet", async (req, res) => {
	const { spreadSheetId, sheetName, lengthRows } = req.body;

	try {
		if (!spreadSheetId) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required spreadSheetId");
		}

		if (!sheetName) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required sheetName");
		}
		
		if (!lengthRows) {
			return throwError(STATUS_CODES.BAD_REQUEST, "Required lengthRows");
		}

		if(isNaN(lengthRows)){
			return throwError(STATUS_CODES.BAD_REQUEST, "lengthRows must be a number");
		}

		const isExist = await checkSheetExistence(spreadSheetId, sheetName);
		console.log("%c  === ","color:cyan","  isExist", isExist)
		if (!isExist) {
			return throwError(STATUS_CODES.NOT_FOUND, `Sheet name: ${sheetName} not found`,);
		}

		const _lengthRows = await getLengthRowsInSheet(spreadSheetId, sheetName);
		if (_lengthRows < lengthRows) {
			return throwError(STATUS_CODES.NOT_FOUND, `Rows length have ${_lengthRows}, not enough`);
		}

		return sendResponse(res, STATUS_CODES.SUCCESS, "Success", isExist);

	} catch (error) {
		if (error.throw) {
			return sendResponse(res, error.status, error.message, false);
		}
		return sendResponse(res, error?.response?.status || 404, error.message, false);
	}
});

/**
 * @swagger
 * /getRows:
 *   get:
 *     summary: Get rows from a Google Spreadsheet
 *     description: Retrieve rows from a specified sheet in a Google Spreadsheet.
 *     parameters:
 *       - in: query
 *         name: spreadSheetId
 *         required: true
 *         description: The ID of the Google Spreadsheet.
 *         example: "1azJcG6h1i-cvLI5pNAmeozjYO5-rX4RfprVcKdYN7_Q"
 *         schema:
 *           type: string
 *       - in: query
 *         name: sheetName
 *         required: true
 *         description: The name of the sheet to retrieve rows from.
 *         example: "investor"
 *         schema:
 *           type: string
 *     responses:
 *       200:
 *         description: Successful response
 *         content:
 *           application/json:
 *             example: { rows: [...] }
 *       400:
 *         description: Bad Request
 *         content:
 *           application/json:
 *             example: { status: "Bad Request", message: "Required spreadSheetId or sheetName" }
 *       404:
 *         description: Not Found
 *         content:
 *           application/json:
 *             example: { status: "Not Found", message: "Sheets not found" }
 *       500:
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             example: { status: "Internal Server Error", message: "An error occurred" }
 */
googleSheetRouter.get("/getRows", async (req, res) => {
	try {
	  const { spreadSheetId, sheetName } = req.query;
	  if (!spreadSheetId) {
		return sendResponse(res, STATUS_CODES.BAD_REQUEST, "Required spreadSheetId", null);
	  }
	  if (!sheetName) {
		return sendResponse(res, STATUS_CODES.BAD_REQUEST, "Required sheetName", null);
	  }
  
	  const isExist = await checkSheetExistence(spreadSheetId, sheetName);
	  if (!isExist) {
		throw new Error("Sheets not found");
	  }
  
	  const response = await getRowsFromSheet(spreadSheetId, sheetName);
	  return sendResponse(res, STATUS_CODES.SUCCESS, "Success", response);
  
	} catch (error) {
	  return sendResponse(res, error?.response?.status || STATUS_CODES.INTERNAL_SERVER_ERROR, error.message, null);
	}
  });

  
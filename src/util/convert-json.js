import path from "path";
import * as fs from "fs/promises";

const currentModuleUrl = new URL(import.meta.url);
const currentModuleDir = path.dirname(currentModuleUrl.pathname);
const projectDir = path.join(currentModuleDir, "../..");
const credentialDir = path.join(currentModuleDir, "..");

async function convertObjectToJSONFile(object, fileName) {
	try {
		if (!object) throw Error("Required data object to convert!");
		if (!fileName) throw Error("Required fileName!");

		var json = JSON.stringify(object);
		await fs.writeFile(projectDir + `/temp/${fileName}.json`, json, "utf8", () => {});
	} catch (error) {
		throw error;
	}
}

async function convertJSONFileToObject(fileName) {
	try {
		if (!fileName) throw Error("Required fileName!");
		const object = await fs.readFile(projectDir + `/temp/${fileName}.json`, "utf8", function readFileCallback(err, data) {
			if (err) {
				throw err
			} else {
				const object = JSON.parse(data); //now it an object
				return object;
			}
		});
		return object
	} catch (error) {
		throw error;
	}
}

async function convertObjectCredentialToJSONFile(object, fileName) {
	try {
		console.log("%c  === ","color:cyan","  credentialDir", credentialDir)

		if (!object) throw Error("Required data object to convert!");
		if (!fileName) throw Error("Required fileName!");

		var json = JSON.stringify(object);
		await fs.writeFile(credentialDir + `/config/${fileName}.json`, json, "utf8", () => {});
	} catch (error) {
		throw error;
	}
}


export { convertObjectToJSONFile, convertJSONFileToObject, convertObjectCredentialToJSONFile};

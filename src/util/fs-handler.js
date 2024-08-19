import fs from "fs";
import path from "path";

async function deleteFilesInFolderAsync(folderPath) {
	if (fs.existsSync(folderPath)) {
		const files = fs.readdirSync(folderPath);

		for (const file of files) {
			const currentPath = path.join(folderPath, file);
			if (fs.lstatSync(currentPath).isFile()) {
				// Delete file
				await fs.promises.unlink(currentPath);
				console.log(`File deleted: ${currentPath}`);
			}
		}
	} else {
		console.log("no files in folder docs");
	}
}

async function deleteFilesAsync(filePath) {
	if (fs.existsSync(filePath)) {
		await fs.promises.unlink(filePath);
		console.log(`File deleted: ${filePath}`);
	} else {
		console.log("no files in folder docs");
	}
}

export { deleteFilesInFolderAsync, deleteFilesAsync };

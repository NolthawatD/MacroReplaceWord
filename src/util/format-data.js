// #remove topic 
async function columnFormatData(columnFormat = "A", rowsData) {
	console.log("%c  === ", "color:cyan", "  columnFormat", columnFormat);
	console.log("%c  === ", "color:cyan", "  rowsData", rowsData);

	const rowIndex = rowsData[0].findIndex((element) => element === columnFormat);

	// const rowIndex = ALPHABET.indexOf(columnFormat);
	if (rowIndex !== -1) {
		console.log(`columnFormat of ${columnFormat} is:`, rowIndex);

		// #check rowsdata < rowIndex
		if (rowIndex > rowsData[0].length) {
			console.log("rowIndex > rows or mismatch");
		}

		const $rowsData = rowsData.slice(1, rowsData.length + 1); // -> ถ้าไม่ได้กำหนด จำนวนแถวมาให้เอาจน index สุดท้าย

		// formatdata to new formatdata
		const formatRowsData = await Promise.all(
			$rowsData.map(async (rowData) => {
				const fomatData = await formatDataIDCard(rowData[rowIndex]); // - - >  format data
				console.log("%c  === ", "color:cyan", "  fomatData", fomatData);
				rowData[rowIndex] = fomatData;
				return [...rowData];
			})
		);

		// console.log("%c  === ", "color:cyan", "  formatRowsData", formatRowsData);
		return formatRowsData;
	} else {
		console.log(`${columnFormat} not found in the array.`);
	}
}

async function calculateRange(data) {
	const numRows = data.length;
	const numColumns = data[0].length;

	// Convert column index to A1 notation
	const columnToLetter = (column) => {
		let temp,
			letter = "";
		while (column > 0) {
			temp = (column - 1) % 26;
			letter = String.fromCharCode(temp + 65) + letter;
			column = (column - temp - 1) / 26;
		}
		return letter;
	};

	const startColumnLetter = columnToLetter(1); // Start from column A
	const endColumnLetter = columnToLetter(numColumns);

	// Generate range string
	const range = `${startColumnLetter}1:${endColumnLetter}${numRows}`;
	return range;
}

async function formatDataIDCard(value) {
	value = value.toString();

	// x-xxxx-xxxx1-52-x
	// 0_1234_56789_12_3

	let replacedValue = "";
	for (let i = 0; i < value.length; i++) {
		if (i >= 0 && i <= 8) {
			replacedValue += "x";
			if (i === 0 || i === 4) {
				replacedValue += "-";
			}
		} else {

			if (i === 10 || i === 12) {
				if (i === 10) {
					replacedValue += "-";
					replacedValue += value[i];
				}

				if (i === 12) {
					replacedValue += "-";
					replacedValue += "x";
				}
			} else {
				replacedValue += value[i];
			}
		}
	}

	return replacedValue;
}

export { columnFormatData, calculateRange };

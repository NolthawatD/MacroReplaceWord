
const createArrayObjectToBatchUpdate = async (object) => {
	let requests = [];

	Object.keys(object).forEach((key) => {
		let placeholder = "$" + `{{${key}}}`;
		let value = object[key];

		let request = {
			replaceAllText: {
				containsText: {
					text: placeholder,
					matchCase: true,
				},
				replaceText: value,
			},
		};

		requests.push(request);
	});

	return requests;
};

const createArrayObjectToBatchUpdateDemo1 = async (object) => {
	let requests = [];

	Object.keys(object).forEach((key) => {
		// let placeholder = "$" + `{{${key}}}`;
		let placeholder = `<<[${key}]>>`;
		let value = object[key];

		// If the value is an array, join its elements into a single string
		if (Array.isArray(value)) {
			value = value.join("\n"); // Or any other separator you prefer
		}

		let request = {
			replaceAllText: {
				containsText: {
					text: placeholder,
					matchCase: true,
				},
				replaceText: value,
			},
		};

		requests.push(request);
	});

	return requests;
};

export { createArrayObjectToBatchUpdate, createArrayObjectToBatchUpdateDemo1}
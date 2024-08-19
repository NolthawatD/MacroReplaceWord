function convertStringToDate(dateString) {
	try {
		const parts = dateString.split("/");
		if (parts.length !== 3) {
			throw new Error("Invalid date format. Please use the format dd/mm/yyyy");
		}
	
		const day = parseInt(parts[0], 10);
		const month = parseInt(parts[1], 10) - 1; // Month is zero-based in Date object
		const year = parseInt(parts[2], 10);
	
		return new Date(year, month, day);
	} catch (error) {
		throw error
	}
}

function formatDateThaiBuddhist(date) {

	const monthsThai = [
		"มกราคม",
		"กุมภาพันธ์",
		"มีนาคม",
		"เมษายน",
		"พฤษภาคม",
		"มิถุนายน",
		"กรกฎาคม",
		"สิงหาคม",
		"กันยายน",
		"ตุลาคม",
		"พฤศจิกายน",
		"ธันวาคม",
	];

	const yearThai = date.getFullYear() + 0; // +543 Convert to Thai Buddhist year
	const monthThai = monthsThai[date.getMonth()];
	return `${monthThai} พ.ศ.${yearThai}`;
}

export { convertStringToDate, formatDateThaiBuddhist };

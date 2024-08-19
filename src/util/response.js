function sendResponse(res, status = 404, message, result = null) {
	const response = {
		status,
		message,
		result,
	};
	return res.status(status).json(response);
}

export { sendResponse };

function throwError(status = 404, msgError) {
	throw {
		throw: true,
		status,
		message : msgError,
	};
}

export { throwError };

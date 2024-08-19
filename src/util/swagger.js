import swaggerJsdoc from "swagger-jsdoc";

const options = {
	explorer: true,
	swaggerDefinition: {
		openapi: "3.0.0",
		info: {
			title: "REST API Docs",
			version: "1.0.0",
			description: "API description",
		},
		components: {
			securitySchemes: {
				bearerAuth: {
					type: "http",
					scheme: "bearer",
					bearerFormat: "JWT",
				},
			},
		},
		security: [
			{
				bearerAuth: [],
			},
		],
		servers: [
			{
				url: "http://localhost:8080/api/v1",
			},
		],
	},
	apis: ["./src/*.js", "./src/**/*.js"],
};

const specs = swaggerJsdoc(options);

export { specs };

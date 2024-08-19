# Use Node.js 18.18.2 as base image
FROM node:18-alpine

# Set working directory in the container
WORKDIR /app

# Copy package.json and package-lock.json to container
COPY package*.json ./

# Install dependencies
RUN npm install 

# Copy the rest of the application code to container
COPY . .

# Expose port 3000 (assuming your application runs on this port)
EXPOSE 8080

# Set environment variables
ENV PORT=8080

# Command to run the application
CMD ["npm", "start"]

# Create a volume to persist data
# VOLUME /app/config
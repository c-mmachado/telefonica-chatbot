import { Request, Response, Router } from "express";
import jwt from "jsonwebtoken";

import { config } from "../../../config/config";
import { compareSync, hashSync } from "bcryptjs";

export const router = Router();

// Define the token endpoint
router.post("/", async (req: Request, res: Response): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    // Check if the request body is empty
    if (!req.body) {
        // If the request body is empty, log a warning and return 400
        console.warn(`Request body is empty`);
        res.status(400).json({
            status: 400,
            statusText: "Bad Request",
            message: "Missing request body",
            data: {
                error: {
                    type: "MissingRequestBody",
                    message: "Request body cannot be empty",
                },
            },
        });
        return;
    }

    // Check if the request body is a valid JSON object
    if (typeof req.body !== "object" || Object.keys(req.body).length === 0) {
        // If the request body is not a valid JSON object, log a warning and return 400
        console.warn(`Request body is not a valid JSON object`);
        res.status(400).json({
            status: 400,
            statusText: "Bad Request",
            message: "Invalid request body",
            data: {
                error: {
                    type: "InvalidRequestBody",
                    message: "Request body must be a valid JSON object",
                    fields: {
                        username: { in: "body", required: true, type: "string" },
                        password: { in: "body", required: true, type: "string" },
                    },
                },
            },
        });
        return;
    }

    // Check if the request body contains a valid username and password fields
    if (
        !req.body.hasOwnProperty("username") ||
        !req.body.hasOwnProperty("password") ||
        !req.body.username ||
        !req.body.password ||
        typeof req.body.username !== "string" ||
        typeof req.body.password !== "string"
    ) {
        // If username or password is not provided, log a warning and return 400
        console.warn(`Missing username or password in request body`);
        const missingFields: any = {};
        if (!req.body.username) {
            missingFields["username"] = { in: "body", required: true, type: "string" };
        }
        if (!req.body.password) {
            missingFields["password"] = { in: "body", required: true, type: "string" };
        }
        res.status(400).json({
            status: 400,
            statusText: "Bad Request",
            message: "Username and password are required",
            data: {
                error: {
                    type: "InvalidRequestBody",
                    message: `Field(s) ${!req.body.username ? "'username' " : ""}${
                        Object.keys(missingFields).length > 1 ? "and " : ""
                    }${!req.body.password ? "'password'" : ""} ${
                        Object.keys(missingFields).length > 1 ? "are" : "is"
                    } required in the request body`,
                    fields: missingFields,
                },
            },
        });
        return;
    }

    // Retrieve the username and password from the request body
    const { username, password } = req.body;
    if (!username || !password || typeof username !== "string" || typeof password !== "string") {
        // If username or password is not provided, log a warning and return 400
        console.warn(`Missing username or password in request body`);
        const missingFields: any = {};
        if (!req.body.username) {
            missingFields["username"] = { in: "body", required: true, type: "string" };
        }
        if (!req.body.password) {
            missingFields["password"] = { in: "body", required: true, type: "string" };
        }
        res.status(400).json({
            status: 400,
            statusText: "Bad Request",
            message: "Username and password are required",
            data: {
                error: {
                    type: "InvalidRequestBody",
                    message: `Field(s) ${!req.body.username ? "'username' " : ""}${
                        Object.keys(missingFields).length > 1 ? "and " : ""
                    }${!req.body.password ? "'password'" : ""} ${
                        Object.keys(missingFields).length > 1 ? "are" : "is"
                    } required in the request body`,
                    fields: missingFields,
                },
            },
        });
        return;
    }

    // Log the received username and password (for debugging purposes, avoid logging sensitive information in production)
    console.debug(`Received username: '${username}', password: '${password}'`);

    // Hardcode the expected username and password for demonstration purposes
    const expectedUsername = config.server.jwt.rootUsername;
    const expectedPassword = config.server.jwt.rootPassword;

    // Hash the expected password for comparison since the password would typically be stored hashed in a database
    const expectedPasswordHash = hashSync(expectedPassword);

    console.debug(`Expected username: '${expectedUsername}', expected password hash: '${expectedPasswordHash}'`);

    // Check if the provided username and password match the expected values
    if (username === expectedUsername && compareSync(password, expectedPasswordHash)) {
        // If the credentials are valid, generate a token
        const token = jwt.sign({ username: expectedUsername }, config.server.jwt.secret, {
            expiresIn: "1h", // Token expiration time
        });
        console.debug(`Token generated successfully`);

        // Send the token in the response
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Token generated successfully",
            data: { token },
        });
    } else {
        // If the credentials are invalid, log a warning and return 401
        console.warn(`Invalid username or password`);
        res.status(401).json({
            status: 401,
            statusText: "Unauthorized",
            message: "Invalid username or password",
            data: {
                error: {
                    type: "InvalidCredentials",
                    message: "Invalid username or password",
                },
            },
        });
    }
});

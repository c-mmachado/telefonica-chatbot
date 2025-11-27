import { Router, Response, Request, NextFunction } from "express";

import { logsRepository as repository } from "../../../config/db";

export const router = Router();

router.get("/", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        // Fetch the api logs from the database
        const result = await repository.logs();

        // Send the api logs as a JSON response
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Logs fetched successfully",
            data: { logs: result },
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while fetching logs",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

router.post("/", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        if (typeof req?.body?.message !== "string") {
            // If the request body does not contain a message field, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Field 'message' is required in the request body",
                data: {
                    error: {
                        type: "InvalidRequestBody",
                        fields: {
                            message: {
                                message: "Field 'message' is required in the request body",
                                in: "body",
                                required: true,
                                type: "string",
                            },
                        },
                    },
                },
            });
            return;
        }

        // Creates a new api log in the database
        const result = await repository.createLog(req.body.message);

        // Send the api log creation result as a JSON response
        res.status(201).json({
            status: 201,
            statusText: "Created",
            message: "Log created successfully",
            data: { log: result },
        });
    } catch (error: any) {
        console.error(`error:`, error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while creating the log",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

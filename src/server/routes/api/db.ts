import { Router, Response, Request, NextFunction } from "express";

import { dbConnection } from "../../../config/db";

export const router = Router();

// Database health check endpoint to verify that the database is running
router.get("/health", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        // Connect to the database
        await dbConnection.connect();

        console.debug(`Connected to database`);

        // Return a 200 status code to indicate that the database is running
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Database connection successful",
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;

        // Return a 503 status code to indicate that the database connection failed
        res.status(503).send({
            status: 503,
            statusText: "Service Unavailable",
            message: "Database connection failed",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    } finally {
        // Close the database connection after the request is complete
        await dbConnection.close();
    }

    next();
});

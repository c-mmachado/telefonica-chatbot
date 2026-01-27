import { Router, Response, Request, NextFunction } from "express";

import { rt } from "../../../utils/client/rt/client";

export const router = Router();

// Database health check endpoint to verify that the database is running
router.get("/health", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        const response = await rt.rt.request.get();

        // Return a 200 status code to indicate that the API connection was successful
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "RT API connection successful",
            data: { rt: response },
        });
    } catch (error: any) {
        console.error(error);

        res.status(503).json({
            status: 503,
            statusText: "Service Unavailable",
            message: "RT API connection failed",
            data: {
                error: {
                    ...error,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

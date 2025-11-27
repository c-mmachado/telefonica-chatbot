import { Router, Response, Request, NextFunction } from "express";
import { AuthenticationResult } from "@azure/msal-node";

import { msalClient } from "../../../utils/client/graph";
import { DefaultSharepointClient } from "../../../utils/client/sharepoint";

export const router = Router();

// Microsoft Sharepoint connection check endpoint to verify that the API is running
router.get("/health", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    // Attempt to connect to the Graph API using the graphClient instance with a me request
    await msalClient
        .acquireTokenByClientCredential({
            scopes: [DefaultSharepointClient.DefaultScope],
        })
        .then((response: AuthenticationResult | null): void => {
            // Return a 200 status code to indicate that the API is running
            res.status(200).json({
                status: 200,
                statusText: "OK",
                message: "SharePoint API connection successful",
                data: { sharepoint: response },
            });
        })
        .catch((error: any) => {
            console.error(error);

            const { body, ...rest } = error;

            // Return a 503 status code to indicate that the API is not running
            res.status(503).json({
                status: 503,
                statusText: "Service Unavailable",
                message: "SharePoint API connection failed",
                data: {
                    error: {
                        ...rest,
                        type: error.name,
                        message: error.message,
                    },
                },
            });
        });
    next();
});

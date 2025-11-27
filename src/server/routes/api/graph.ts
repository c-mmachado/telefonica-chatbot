import { Router, Response, Request, NextFunction } from "express";
import { AuthenticationResult } from "@azure/msal-node";

import { DefaultGraphClient, msalClient } from "../../../utils/client/graph";

export const router = Router();

// Microsoft Graph cconnection check endpoint to verify that the API is running
router.get("/health", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    // Attempt to connect to the Graph API using the graphClient instance with a me request
    await msalClient
        .acquireTokenByClientCredential({
            scopes: [DefaultGraphClient.DefaultScope],
        })
        .then((response: AuthenticationResult | null) => {
            if (response instanceof Error) {
                throw response;
            }

            // Return a 200 status code to indicate that the API is running
            res.status(200).json({
                status: 200,
                statusText: "OK",
                message: "Microsoft Graph API connection successful",
                data: { result: response },
            });
        })
        .catch((error: any) => {
            console.error(error);

            const { body, ...rest } = error;

            // Return a 503 status code to indicate that the API is not running
            res.status(503).json({
                status: 503,
                statusText: "Service Unavailable",
                message: "Unable to connect to Microsoft Graph API",
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

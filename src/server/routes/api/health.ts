import { Router, Request, Response, NextFunction } from "express";

export const router: Router = Router();

// Health check endpoint for the express app to verify that the API is running
router.get("/", async (req: Request, res: Response, _: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    // Return a 200 status code to indicate that the API is running
    res.status(200).json({ status: 200, statusText: "OK", message: "API is running" });
});

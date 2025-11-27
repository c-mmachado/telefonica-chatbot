import { TurnContext } from "botbuilder";
import { Router, Response, Request, NextFunction } from "express";

import { adapter } from "../../../config/adapter";
import { bot } from "../../../bots/initialize";

export const router = Router();

// Register an API endpoint with `express`. Teams sends messages to your application
// through this endpoint.
//
// The Teams Toolkit bot registration configures the bot with `/api/messages` as the
// Bot Framework endpoint. If you customize this route, update the Bot registration
// in `infra/botRegistration/azurebot.bicep`.
// Process Teams activity with Bot Framework.
router.post("/", async (req: Request, res: Response, _: NextFunction): Promise<void> => {
    await adapter
        .process(req, res, async (context: TurnContext): Promise<any> => {
            console.debug(`req.headers:`, req.headers);
            return await bot.run(context);
        })
        .catch((error: any) => {
            console.error(error);

            if (!error.message?.includes("412")) {
                // Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, shouldn't throw this error
                throw error;
            }
        });
});

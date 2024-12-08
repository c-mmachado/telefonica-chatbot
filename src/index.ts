// Import required packages
import { MemoryStorage, TurnContext } from "botbuilder";
import express from "express";
import path from "path";
import send from "send";
import "isomorphic-fetch";

import {
  commandBot,
  conversationState,
  userState,
} from "./config/initialize";
import { SimpleConversationReferenceStore, TeamsBot } from "./bots/teamsBot";

// Define a simple conversation reference store.
const conversationReferences: SimpleConversationReferenceStore = {};

// Create the activity handler.
const bot = new TeamsBot(
  conversationState,
  userState,
  new MemoryStorage(),
  conversationReferences
);

// Create express application.
const expressApp = express();
expressApp.use(express.json());

const server = expressApp.listen(
  process.env.port || process.env.PORT || 3978,
  () => {
    console.log(
      `[expressApp][INFO] Bot started, ${expressApp.name} listening to`,
      server.address()
    );
  }
);

// Register an API endpoint with `express`. Teams sends messages to your application
// through this endpoint.
//
// The Teams Toolkit bot registration configures the bot with `/api/messages` as the
// Bot Framework endpoint. If you customize this route, update the Bot registration
// in `templates/azure/provision/botservice.bicep`.
// Process Teams activity with Bot Framework.
expressApp.post("/api/messages", async (req, res) => {
  await commandBot
    .requestHandler(req, res, async (context: TurnContext): Promise<any> => {
      console.debug(`[expressApp][DEBUG] [${req.method}] req.url: ${req.url}`);
      console.debug(
        `[expressApp][DEBUG] [${req.method}] req.headers:\n${JSON.stringify(
          req.headers,
          null,
          2
        )}`
      );
      return await bot.run(context);
    })
    .catch((err) => {
      console.error(
        `[expressApp][ERROR] [${req.method}] error:\n${JSON.stringify(
          err,
          null,
          2
        )}`
      );

      // Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, shouldn't throw this error.
      if (!err.message.includes("412")) {
        throw err;
      }
    });
});

// Allow the auth-start.html and auth-end.html to be served from the public folder.
// expressApp.get(["/auth-start.html", "/auth-end.html"], async (req, res) => {
//   console.debug(`[expressApp][DEBUG] [${req.method}] req.url: ${req.url}`);
//   console.debug(
//     `[expressApp][DEBUG] [${req.method}] req.originalUrl:\n${JSON.stringify(req.originalUrl, null, 2)}`
//   );
//
//   send(
//     req,
//     path.join(
//       __dirname,
//       "public",
//       req.url.includes("auth-start.html") ? "auth-start.html" : "auth-end.html"
//     )
//   ).pipe(res);
// });

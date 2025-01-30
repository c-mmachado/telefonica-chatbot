import {
  ConversationState,
  MemoryStorage,
  TurnContext,
  UserState,
} from "botbuilder";
import express, { Response, Request } from "express";
// import path from "path";
// import send from "send";
import "isomorphic-fetch";
import * as mssql from "mssql";

import { commandBot } from "./config/initialize";
import { TeamsBot } from "./bots/teamsBot";
import {
  ConversationReferenceStore,
  DefaultHandlerManager,
  HandlerContextManager,
  HandlerManager,
} from "./commands/handlerManager";
import { TicketCommandHandler } from "./commands/ticket/ticket";
import { config } from "./config/config";
import { APIClient } from "./utils/apiClient";
import { AuthCommandDispatchDialog } from "./dialogs/authCommandDispatchDialog";
import { AuthRefreshActionHandler } from "./adaptiveCards/actions/authRefresh/authRefresh";
import { TicketAdaptiveCardCreateActionHandler } from "./adaptiveCards/actions/ticket/create";
import { TicketAdaptiveCardCancelActionHandler } from "./adaptiveCards/actions/ticket/cancel";

// Define the state store for your bot.
// See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
// A bot requires a state storage system to persist the dialog and user state between messages
const memoryStorage: MemoryStorage = new MemoryStorage();

// Create conversation and user state with the storage provider defined above
export const conversationState: ConversationState = new ConversationState(
  memoryStorage
);
export const userState: UserState = new UserState(memoryStorage);

// Define a simple conversation reference store
const conversationStore: ConversationReferenceStore = {};

// Create the API client to the ticketing API
const apiClient: APIClient = new APIClient(config);

// Create the context manager
const contextManager: HandlerContextManager = new HandlerContextManager(
  config,
  conversationStore
);

// Create the handler manager
const handlerManager: HandlerManager = new DefaultHandlerManager(
  contextManager,
  {
    commands: [new TicketCommandHandler(apiClient)],
    actions: [
      new AuthRefreshActionHandler(),
      new TicketAdaptiveCardCreateActionHandler(config, apiClient),
      new TicketAdaptiveCardCancelActionHandler(),
    ],
  }
);

// Create the auth flow dialog
const dialog: AuthCommandDispatchDialog = new AuthCommandDispatchDialog(
  config,
  conversationState,
  new MemoryStorage(),
  handlerManager
);

// Register the dialog with the context manager
contextManager.registerDialog(dialog);

// Create databse connection
const dbConnection: mssql.ConnectionPool = new mssql.ConnectionPool({
  server: config.dbHost,
  port: config.dbPort,
  user: config.dbUser,
  password: config.dbPassword,
  database: config.dbName,
  options: {
    encrypt: false,
    enableArithAbort: true,
  },
});

// Create the activity handler.
const bot: TeamsBot = new TeamsBot(
  config,
  conversationState,
  userState,
  handlerManager,
  dialog
);

// const tmp = async () => {
//   const api = new APIClient(config);
//   const cookie = await api.login();
//   console.debug(`[index][DEBUG] login: ${JSON.stringify(cookie, null, 2)}`);

//   const queues = await api.queues(cookie);
//   console.debug(`[index][DEBUG] queues: ${JSON.stringify(queues, null, 2)}`);
// };
// tmp();

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
// in `infra/botRegistration/azurebot.bicep`.
// Process Teams activity with Bot Framework.
expressApp.post("/api/messages", async (req: Request, res: Response) => {
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
    .catch((err: any) => {
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

// Health check endpoint for the express app to verify that the app is running.
expressApp.get("/health", async (req: Request, res: Response) => {
  console.debug(`[expressApp][DEBUG] [${req.method}] req.url: ${req.url}`);
  console.debug(
    `[expressApp][DEBUG] [${req.method}] req.headers:\n${JSON.stringify(
      req.headers,
      null,
      2
    )}`
  );

  res
    .status(200)
    .send(
      JSON.stringify(
        { status: 200, data: { message: "Bot is running" } },
        null,
        2
      )
    );
});

// Database health check endpoint to verify that the database is running.
expressApp.get("/db/health", async (req: Request, res: Response) => {
  console.debug(`[expressApp][DEBUG] [${req.method}] req.url: ${req.url}`);
  console.debug(
    `[expressApp][DEBUG] [${req.method}] req.headers:\n${JSON.stringify(
      req.headers,
      null,
      2
    )}`
  );

  try {
    await dbConnection.connect();

    console.debug(`[expressApp][DEBUG] [${req.method}] Connected to database`);

    res.status(200).send(
      JSON.stringify(
        {
          status: 200,
          data: { message: "Database connection successful" },
        },
        null,
        2
      )
    );
  } catch (error: any) {
    console.error(
      `[expressApp][ERROR] [${req.method}] error:\n${JSON.stringify(
        error,
        null,
        2
      )}`
    );

    res.status(500).send(
      JSON.stringify(
        {
          data: { status: 500, message: "Database connection failed", error },
        },
        null,
        2
      )
    );
  }
});

// Ticketing API health check endpoint to verify that we can connect to the ticketing API.
expressApp.get("/api/health", async (req: Request, res: Response) => {
  console.debug(`[expressApp][DEBUG] [${req.method}] req.url: ${req.url}`);
  console.debug(
    `[expressApp][DEBUG] [${req.method}] req.headers:\n${JSON.stringify(
      req.headers,
      null,
      2
    )}`
  );

  const cookie = await apiClient.login();
  if (cookie) {
    console.debug(
      `[expressApp][DEBUG] [${req.method}] cookie:\n${JSON.stringify(
        cookie,
        null,
        2
      )}`
    );
    res.status(200).send(
      JSON.stringify(
        {
          status: 200,
          data: { cookie, message: "API connection successful" },
        },
        null,
        2
      )
    );
  } else {
    res
      .status(500)
      .send(
        JSON.stringify(
          { status: 500, data: { message: "API connection failed" } },
          null,
          2
        )
      );
  }
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

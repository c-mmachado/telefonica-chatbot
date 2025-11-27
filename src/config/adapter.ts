// import { TeamsAdapter } from "@microsoft/teams-ai";
// import { BotBuilderCloudAdapter } from "@microsoft/teamsfx";

import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationServiceClientCredentialFactory,
  TurnContext,
} from "botbuilder";
// import ConversationBot = BotBuilderCloudAdapter.ConversationBot;

import { config } from "./config";

// The bot credentials to create a `CloudAdapter`.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
export const adapter = new CloudAdapter(
  new ConfigurationBotFrameworkAuthentication(
    {},
    new ConfigurationServiceClientCredentialFactory({
      MicrosoftAppType: config.botType, // config.tenantId ? "SingleTenant" : "MultiTenant",
      MicrosoftAppId: config.botId,
      MicrosoftAppPassword: config.botPassword,
      MicrosoftAppTenantId: config.tenantId,
    })
  )
);

const adapterTurnErrorHandler = adapter.onTurnError;
const onTurnErrorHandler = async (
  context: TurnContext,
  error: any
): Promise<void> => {
  console.error(error);

  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  await adapterTurnErrorHandler?.(context, error);
};
adapter.onTurnError = onTurnErrorHandler;

// new TeamsAdapter(
//   {},
//   new ConfigurationServiceClientCredentialFactory({
//     MicrosoftAppId: config.botId,
//     MicrosoftAppPassword: config.botPassword,
//     MicrosoftAppType: config.botType,
//     MicrosoftAppTenantId: config.tenantId,
//   })
// );

// Create the command bot and register the command handlers for your app.
// You can also use the commandBot.command.registerCommands to register other commands
// if you don't want to register all of them in the constructor
// export const commandBot = new ConversationBot({
//   adapterConfig: {
//     MicrosoftAppId: config.botId,
//     MicrosoftAppPassword: config.botPassword,
//     MicrosoftAppType: config.botType,
//     MicrosoftAppTenantId: config.tenantId,
//   },
//   // See https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk to learn more about ssoConfig
//   ssoConfig: {
//     aad: {
//       scopes: [
//         "User.Read",
//         "ChatMessage.Read",
//         "ChannelMessage.Read.All",
//         "Team.ReadBasic.All",
//         "Channel.ReadBasic.All",
//         "ProfilePhoto.Read.All",
//         "Files.Read.All",
//       ],
//       initiateLoginEndpoint: `https://${config.botDomain}/auth-start.html`,
//       authorityHost: config.authorityHost,
//       clientId: config.clientId,
//       tenantId: config.tenantId,
//       clientSecret: config.clientSecret,
//     },
//     dialog: {
//       // userState: userState,
//       // conversationState: conversationState,
//       dedupStorage: new MemoryStorage(),
//       ssoPromptConfig: {
//         timeout: 900000,
//         endOnInvalidMessage: true,
//       },
//     },
//   },
//   command: {
//     enabled: false,
//     commands: [],
//     ssoCommands: [],
//   },
// });

// // Creates the SSO token exchange middleware.
// // This middleware is used to exchange the SSO token for a user.
// const tokenExchangeMiddleware = new TeamsSSOTokenExchangeMiddleware(
//   memoryStorage,
//   config.connectionName
// );
// commandBot.adapter.use(tokenExchangeMiddleware);

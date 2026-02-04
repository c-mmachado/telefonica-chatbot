import { ConversationState, MemoryStorage, TeamsActivityHandler, UserState } from "botbuilder";

import { DefaultHandlerManager, HandlerManager } from "./commands/manager";
import { DefaultHandlerTurnContextFactory, HandlerTurnContextFactory } from "./commands/context";
import { ConversationHelper, ConversationReferenceStore, DefaultConversationHelper } from "./commands/conversation";
import { DefaultDialogManager, DialogManager } from "./dialogs/manager";
import { OAuthDialog } from "./dialogs/oauthDialog";

import { AuthRefreshActionHandler } from "./adaptiveCards/actions/authRefresh/authRefresh";
import { TicketCommandHandler } from "./commands/ticket/ticket";
import { TicketAdaptiveCardNextActionHandler } from "./adaptiveCards/actions/ticket/next";
import { TicketAdaptiveCardCreateActionHandler } from "./adaptiveCards/actions/ticket/create";
import { TicketAdaptiveCardCancelActionHandler } from "./adaptiveCards/actions/ticket/cancel";
import { TicketAdaptiveCardSelectChoiceActionHandler } from "./adaptiveCards/actions/ticket/selectChoice";

import { config } from "../config/config";
import { techRepository } from "../config/db";

import { rt as rtClient } from "../utils/client/rt/client";
import { graphClient } from "../utils/client/graph";
import { ActivityErrorHandlerFactory, ActivityHandlerFactory } from "./teamsBot";

// Define the state store for your bot.
// See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
// A bot requires a state storage system to persist the dialog and user state between messages
const memoryStorage: MemoryStorage = new MemoryStorage();

// Create conversation and user state with the storage provider defined above
const conversationState: ConversationState = new ConversationState(memoryStorage);
const userState: UserState = new UserState(memoryStorage);

// Define a simple conversation reference store
const conversationStore: ConversationReferenceStore = {};

// Creates the dialog manager
const dialogManager: DialogManager = new DefaultDialogManager();

// Create the context manager
const conversationHelper: ConversationHelper = new DefaultConversationHelper(config, conversationStore);

// Create the commands and actions
const arah = new AuthRefreshActionHandler();

const tch = new TicketCommandHandler(rtClient);
const tnxah = new TicketAdaptiveCardNextActionHandler(rtClient);
const tpah = new TicketAdaptiveCardCreateActionHandler(config, rtClient, graphClient);
const tnah = new TicketAdaptiveCardCancelActionHandler();
const tscah = new TicketAdaptiveCardSelectChoiceActionHandler(rtClient);

const contextFactory: HandlerTurnContextFactory = new DefaultHandlerTurnContextFactory(
    dialogManager,
    conversationHelper,
);

// Create the handler manager
const handlerManager: HandlerManager = new DefaultHandlerManager(userState, config, graphClient, dialogManager, {
    commands: [tch],
    actions: [arah, tnxah, tpah, tnah, tscah],
});

// Create the auth flow dialog and register it in the dialog manager
const dialog: OAuthDialog = new OAuthDialog(config, conversationState, new MemoryStorage());
dialogManager.registerDialog(dialog);

const errorHandler = ActivityErrorHandlerFactory.Default.create();

// Create the activity handler for incoming Microsoft Teams activities
export const bot: TeamsActivityHandler = ActivityHandlerFactory.Default.create({
    config: config,
    conversationState: conversationState,
    userState: userState,
    handlerManager: handlerManager,
    techRepository: techRepository,
    contextFactory: contextFactory,
    errorHandler: errorHandler,
});

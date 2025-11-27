import { ConversationReference, TurnContext } from "botbuilder";
import { DialogTurnResult } from "botbuilder-dialogs";
import { Handler } from "express";
import { HandlerTriggerData } from "src/bots/commands/manager";

// Put whatever you add to the context here
declare module "botbuilder-core" {
  interface TurnContext {
    trigger(): HandlerTriggerData;

    request<T>(): ActiveRequest<T>;

    switchToConversation(
      conversationReference: Partial<ConversationReference>,
      action: (context: TurnContext) => Promise<void>
    ): Promise<void>;

    switchToPersonalConversation(
      action: (context: TurnContext) => Promise<void>
    ): Promise<void>;

    runDialog(
      context: TurnContext,
      dialogName: string,
      data?: any
    ): Promise<DialogTurnResult>;

    continueDialog(
      context: TurnContext,
      dialogName: string
    ): Promise<DialogTurnResult>;

    stopDialog(
      context: TurnContext,
      dialogName: string
    ): Promise<DialogTurnResult>;

    runDialog(
      context: TurnContext,
      dialogName: string,
      data?: any
    ): Promise<DialogTurnResult>;
  }
}

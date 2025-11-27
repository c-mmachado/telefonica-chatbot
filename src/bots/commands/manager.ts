import {
  CardFactory,
  ChannelAccount,
  ChannelInfo,
  ConversationAccount,
  ConversationReference,
  MessageFactory,
  SigninStateVerificationQuery,
  StatePropertyAccessor,
  TeamDetails,
  TeamsChannelAccount,
  TeamsInfo,
  TokenResponse,
  TurnContext,
  UserState,
} from "botbuilder";
import { DialogTurnResult, DialogTurnStatus } from "botbuilder-dialogs";
import { nanoid } from "nanoid";
import { Channel, ChatMessage } from "@microsoft/microsoft-graph-types";

import * as ACData from "adaptivecards-templating";

import { BotConfiguration } from "../../config/config";
import {
  ActionHandler,
  CommandHandler,
  Handler,
  OAuthCommandHandler,
} from "./handler";
import { HandlerMessage } from "./message";
import { DialogManager } from "../dialogs/manager";
import { OAuthDialog } from "../dialogs/oauthDialog";
import {
  AdaptiveCardAction,
  AdaptiveCardActionAuthRefreshDataInput,
} from "../adaptiveCards/actions/actions";
import { GraphClient } from "../../utils/client/graph";

import authRefreshCard from "../adaptiveCards/templates/authRefreshCard.json";

export enum HandlerType {
  Command,
  Action,
}

export interface HandlerManager {
  resolve(
    pattern: string,
    type: HandlerType.Command
  ): [CommandHandler | null, HandlerMessage | null];
  resolve(
    pattern: string,
    type: HandlerType.Action
  ): [ActionHandler | null, HandlerMessage | null];
  resolve(
    pattern: string,
    type: HandlerType
  ): [Handler | null, HandlerMessage | null];

  dispatch(
    command: Handler,
    context: TurnContext,
    message: HandlerMessage,
    data?: any
  ): Promise<any>;

  resolveAndDispatch<T>(
    context: TurnContext,
    message: string,
    data?: any
  ): Promise<T>;

  onSignInAction(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ): Promise<void>;
}

export interface HandlerTriggerData {
  ref: Partial<ConversationReference>;
  team: TeamDetails;
  channel: Channel;
  conversation: ConversationAccount;
  thread: ChatMessage;
  threadFrom: TeamsChannelAccount;
  reply: ChatMessage;
  replyFrom: TeamsChannelAccount;
}

export interface ActiveRequest<T> {
  ref: Partial<ConversationReference>;
  user: ChannelAccount;
  requestId: string;
  timestamp: number;
  expiresAt: number;
  trigger?: Partial<HandlerTriggerData>;
  data?: T;
}

export interface DefaultHandlerManagerOptions {
  commands: CommandHandler[];
  actions: ActionHandler[];
}

export class DefaultHandlerManager implements HandlerManager {
  public static readonly RequestIdKey: Symbol = Symbol("RequestId");
  public static readonly ActiveRequestKey: Symbol = Symbol("ActiveRequest");

  private static readonly ActiveRequestsMapProperty: string =
    "ActiveRequestsMapProperty";

  private readonly _activeRequestsMapAccessor: StatePropertyAccessor<
    Record<string, ActiveRequest<any>>
  >;

  constructor(
    readonly userState: UserState,
    private readonly _config: BotConfiguration,
    private readonly _graph: GraphClient,
    private readonly _dialogs: DialogManager,
    // private readonly _conversations: ConversationHelper,
    private readonly _options: Partial<DefaultHandlerManagerOptions>
  ) {
    this._activeRequestsMapAccessor = userState.createProperty<
      Record<string, ActiveRequest<unknown>>
    >(DefaultHandlerManager.ActiveRequestsMapProperty); // requestId -> request
  }

  public resolve(
    pattern: string,
    type: HandlerType.Command
  ): [CommandHandler | null, HandlerMessage | null];

  public resolve(
    pattern: string,
    type: HandlerType.Action
  ): [ActionHandler | null, HandlerMessage | null];

  public resolve(
    pattern: string,
    type: HandlerType
  ): [Handler | null, HandlerMessage | null] {
    switch (type) {
      case HandlerType.Command:
        return this._resolveFromText(this._options.commands || [], pattern);
      case HandlerType.Action:
        return this._resolveFromText(this._options.actions || [], pattern);
      default:
        return [null, null];
    }
  }

  private _resolveFromText(
    handlers: Handler[],
    pattern: string
  ): [Handler | null, HandlerMessage | null] {
    let message: HandlerMessage | null = null;

    for (const handler of handlers) {
      if (!handler?.pattern) {
        continue;
      }
      if (handler.pattern instanceof RegExp) {
        const matches: RegExpMatchArray | null = pattern.match(handler.pattern);
        if (matches) {
          message = {
            text: pattern,
            matches,
          };
          return [handler, message];
        }
      } else {
        if (handler.pattern === pattern) {
          message = {
            text: pattern,
            matches: undefined,
          };
          return [handler, message];
        }
      }
    }
    return [null, null];
  }

  public async dispatch(
    handler: Handler,
    context: TurnContext,
    message: HandlerMessage,
    data?: any
  ): Promise<any> {
    if (!handler) {
      console.error(`Argument 'handler' must be provided`);
      throw new Error(`Argument 'handler' must be provided`);
    }

    if (!context) {
      console.error(`Argument 'context' must be provided`);
      throw new Error(`Argument 'context' must be provided`);
    }

    console.debug(`handler type: '${handler.constructor.name}'`);
    console.debug("data:", data);

    // Generate or retrieve the requestId for the current handling and store it in the turn state for later retrieval
    const requestId: string = data?.requestId
      ? data.requestId
      : context.activity.channelData?.requestId
      ? context.activity.channelData.requestId
      : nanoid(); // customAlphabet("1234567890abcdefghijklmnopqrstuvwxyz", 24)();
    context.turnState.set(DefaultHandlerManager.RequestIdKey, requestId);

    console.debug(`requestId: '${requestId}'`);

    const reference = TurnContext.getConversationReference(context.activity);

    const activeRequests: Record<
      string,
      ActiveRequest<unknown>
    > = await this._activeRequestsMapAccessor.get(context, {});

    if (!activeRequests[requestId]) {
      let team: TeamDetails | undefined = undefined;
      let channel: Channel | undefined = undefined;
      let thread: ChatMessage | undefined = undefined;
      let threadFrom: TeamsChannelAccount | undefined = undefined;
      let reply: ChatMessage | undefined = undefined;
      const conversation: ConversationAccount = context.activity.conversation;

      if (conversation.isGroup) {
        // If the conversation is a group conversation, retrieve the team and channel information
        team = await TeamsInfo.getTeamDetails(context).catch(
          (error: any): undefined => {
            console.error(error);
            return undefined;
          }
        );

        // Resolves the channel information where the command was triggered
        const channels: ChannelInfo[] = (
          await TeamsInfo.getTeamChannels(context).catch(
            (error: any): ChannelInfo[] => {
              console.error(error);
              return [];
            }
          )
        ).filter(
          (info: ChannelInfo) =>
            info.id ===
            (conversation.id?.indexOf(";") >= 0
              ? conversation.id.split(";")[0]
              : conversation.id)
        );
        const channelInfo: ChannelInfo | null =
          channels?.length > 0 ? channels[0] : null;

        // Conversation id format: 19:<channelId>;messageid=<messageId>
        let messageId: string =
          conversation.id.indexOf(";") >= 0
            ? conversation.id.split(";")[1]
            : conversation.id;
        messageId = messageId.replace("messageid=", "");

        channel =
          team?.aadGroupId && channelInfo?.id
            ? await this._graph.teams
                .id(team?.aadGroupId)
                .channels.id(channelInfo?.id)
                .get()
                .catch((error: any): undefined => {
                  console.error(error);
                  return undefined;
                })
            : undefined;
        console.debug("Graph channel:", channel);

        thread =
          team?.aadGroupId && channelInfo?.id && messageId
            ? await this._graph.teams
                .id(team?.aadGroupId)
                .channels.id(channelInfo?.id)
                .messages.id(messageId)
                .get()
                .catch((error: any): undefined => {
                  console.error(error);
                  return undefined;
                })
            : undefined;
        console.debug("Graph message:", thread);

        threadFrom = await TeamsInfo.getMember(
          context,
          thread?.from?.user?.id!
        ).catch((error: any): undefined => {
          console.error(error);
          return undefined;
        });

        // TODO: Reply might be in a personal chat, need to handle that case
        reply =
          team?.aadGroupId &&
          channelInfo?.id &&
          messageId &&
          context.activity.id
            ? await this._graph.teams
                .id(team?.aadGroupId)
                .channels.id(channelInfo?.id)
                .messages.id(messageId)
                .replies.id(context.activity.id)
                .get()
                .catch((error: any): undefined => {
                  console.error(error);
                  return undefined;
                })
            : undefined;
      }

      const replyFrom: TeamsChannelAccount | undefined =
        await TeamsInfo.getMember(context, context.activity.from.id).catch(
          (error: any): undefined => {
            console.error(error);
            return undefined;
          }
        );

      // If there is no active request for the current requestId, create one
      activeRequests[requestId] = {
        ref: reference,
        requestId: requestId,
        user: context.activity.from,
        timestamp: Date.now(),
        expiresAt: Date.now() + this._config.botTimeout * 1000,
        trigger: {
          ref: TurnContext.getConversationReference(context.activity),
          team: team,
          channel: channel,
          thread: thread,
          threadFrom: threadFrom,
          reply: reply,
          replyFrom: replyFrom,
        },
      };
      await this._activeRequestsMapAccessor.set(context, activeRequests);
    } else {
      // Update the conversation in case it has changed
      activeRequests[requestId].ref = reference;
    }

    // Store the active request in the turn state for later retrieval
    context.turnState.set(
      DefaultHandlerManager.ActiveRequestKey,
      activeRequests[requestId]
    );

    if (handler instanceof OAuthCommandHandler) {
      return await this._beginOAuthFlow(context, message);
    }
    return await handler.run(context, message);
  }

  public async resolveAndDispatch(
    context: TurnContext,
    message: string,
    data?: any
  ): Promise<any> {
    let handler: Handler | null = null;
    let handlerMessage: HandlerMessage | null = null;
    if (context.activity.name == AdaptiveCardAction.Name) {
      // If the activity name is "adaptiveCard/action", it is an action handler
      [handler, handlerMessage] = this.resolve(message, HandlerType.Action);
    } else {
      // Otherwise, it is a command handler
      [handler, handlerMessage] = this.resolve(message, HandlerType.Command);
    }

    if (!handler) {
      return;
    }

    return await this.dispatch(handler, context, handlerMessage!, data);
  }

  public async onSignInAction(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ): Promise<void> {
    // This activity type can be triggered during the auth flow in either a 'signin/verifyState' or 'signin/tokenExchange' event
    console.debug("query:", query);

    if (!context.activity?.replyToId) {
      console.error(`'context.activity.replyToId' is`, undefined);
      throw new Error(
        "Unable to process signin action: 'context.activity.replyToId' is undefined"
      );
    }

    // Deletes the message corresponding to the auth flow card sent by the bot
    if (context.activity?.replyToId) {
      await context.deleteActivity(context.activity.replyToId);
    }

    // Checks if the auth flow was canceled by the user or completed
    const state = query.state;
    if (typeof state === "string" && state.indexOf("CancelledByUser") >= 0) {
      // If the auth flow was canceled by the user, ends the dialog
      console.debug("User cancelled the authentication flow");
      await context.sendActivity(
        "El usuario rechazó el flujo de autenticación."
      );
      await this._dialogs
        .stopDialog(context, OAuthDialog.name)
        .catch((error: any): void => {
          console.error(error);
        });
    } else {
      // const userTokenClient: UserTokenClient = context.turnState.get(
      //   (context.adapter as CloudAdapter).UserTokenClientKey
      // );
      // if (!userTokenClient) {
      //   throw new Error("'UserTokenClient' not available in this context.");
      // }
      // const tokenResponse = await userTokenClient.getUserToken(
      //   context.activity.from.id,
      //   this._config.botConnectionName,
      //   context.activity.channelId,
      //   state!
      // );
      // console.debug("tokenResponse:", tokenResponse);

      // If the auth flow was completed, continues the dialog to run the next step
      const dialogTurn: DialogTurnResult<{
        requestId?: string;
        tokenResult?: Partial<TokenResponse>;
      }> | null = await this._dialogs
        .continueDialog(context, OAuthDialog.name)
        .catch((error: any): null => {
          console.error(error);
          return null;
        });
      console.debug("Dialog continued with result:", dialogTurn);

      if (
        dialogTurn?.status !== DialogTurnStatus.complete ||
        !dialogTurn?.result
      ) {
        console.error(`Invalid dialog turn result:`, dialogTurn);
        throw new Error(
          `OAuth dialog continuation returned an invalid dialog turn result`
        );
      }

      if (dialogTurn?.result.requestId) {
        const activeRequestsMap: Record<
          string,
          ActiveRequest<{ message: string }>
        > = await this._activeRequestsMapAccessor.get(context, {});

        const activeRequest: ActiveRequest<{ message?: string }> =
          activeRequestsMap[dialogTurn?.result.requestId];

        context.turnState.set(
          DefaultHandlerManager.RequestIdKey,
          dialogTurn?.result.requestId
        );
        context.turnState.set(
          DefaultHandlerManager.ActiveRequestKey,
          activeRequest
        );

        const message: string = activeRequest.data?.message ?? "";
        let handler: Handler | null = null;
        let handlerMessage: HandlerMessage | null = null;
        [handler, handlerMessage] = this.resolve(message, HandlerType.Command);
        if (!handler) {
          console.error(`No handler found for message: '${message}'`);
          throw new Error(`No handler found for message: '${message}'`);
        }

        if (!handlerMessage) {
          console.error(
            `Unable to build handler message object for message: '${message}'`
          );
          throw new Error(
            `Unable to build handler message object for message: '${message}'`
          );
        }

        return await handler.run(
          context,
          handlerMessage,
          dialogTurn.result?.tokenResult?.token
        );
      }
    }
  }

  private async _beginOAuthFlow(
    context: TurnContext,
    message: HandlerMessage
  ): Promise<void> {
    const activeRequest: ActiveRequest<{ message: string }> =
      context.turnState.get(DefaultHandlerManager.ActiveRequestKey);

    activeRequest.data = {
      ...(activeRequest.data || {}),
      message: message.text,
    };
    console.debug(`activeRequest:`, activeRequest);

    // Switch to personal context and send the auth refresh adaptive card
    await (context as any).switchToPersonalConversation(
      async (chatContext: TurnContext): Promise<void> => {
        // A newly created conversation context does not have a value set for the "from" field
        chatContext.activity.from =
          activeRequest.trigger?.replyFrom || chatContext.activity.from;

        const userIds: string[] = [];
        if (activeRequest.trigger?.replyFrom?.id) {
          userIds.push(activeRequest.trigger.replyFrom.id);
        }

        // Creates the card data to be passed to the adaptive card template which will be used to expand the adaptive card template with the actual data
        const cardData: AdaptiveCardActionAuthRefreshDataInput = {
          requestId: activeRequest.requestId,
          userIds: userIds,
        };

        // Expands the adaptive card template with the card data by replacing the placeholders
        // with the actual data
        const cardJson = new ACData.Template(authRefreshCard).expand({
          $root: cardData,
        });

        // oauthHandlerState.lastActivity = wrapper.context.activity;

        await chatContext.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(cardJson))
        );

        console.debug(
          `Sent auth refresh adaptive card to user '${chatContext.activity.from.name}' with id '${chatContext.activity.from.id}' relating to request id '${activeRequest.requestId}' in personal context`
        );

        // await OAuthPrompt.sendOAuthCard(
        //   {
        //     connectionName: this._config.botConnectionName,
        //     text: "Revise y acepte el flujo de consentimiento para continuar.",
        //     title: "Flujo Consentimiento",
        //     timeout: 900000,
        //   },
        //   chatContext
        // );

        // await this._dialogs.runDialog(chatContext, OAuthDialog.name, {
        //   requestId: activeRequest.requestId,
        // });
      }
    );
  }
}

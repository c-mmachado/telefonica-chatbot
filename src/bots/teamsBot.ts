import {
  TeamsActivityHandler,
  TurnContext,
  TeamsInfo,
  TeamsPagedMembersResult,
  Activity,
  ConversationReference,
  ConversationParameters,
  ConversationState,
  UserState,
  CardFactory,
  MessageFactory,
  ConversationAccount,
  TeamDetails,
  SigninStateVerificationQuery,
  ChannelInfo,
  Storage,
} from "botbuilder";
import * as ACData from "adaptivecards-templating";

import {
  SimpleGraphClient,
  TeamChannelMessage,
  TeamChannelMessageAttachment,
} from "../utils/graphClient";
import { AppInstallUtils } from "../utils/appInstall";
import { AuthCommandDispatchDialog } from "../dialogs/authCommandDispatchDialog";
import { adaptiveCardActionName } from "../utils/actions";
import ticketCard from "../adaptiveCards/templates/createTicketCard.json";
import authRefreshCard from "../adaptiveCards/templates/authRefreshCard.json";
import config from "../config/config";

export type SimpleConversationReferenceStore = {
  [key: string]: Partial<ConversationReference>;
};

type AdaptiveCardActionActivityValue = {
  action: {
    verb: string;
    data?: any & {
      command: string;
    };
  };
};

type AdaptiveCardActionActivityValueDataCreateTicket = {
  team: TeamDetails;
  channel: { id: string; name: string };
  conversation: { id: string; name: string };
  from: ConversationAccount;
  token: string;
  ticketStateChoiceSet: string;
  ticketCategoryChoiceSet: string;
  ticketDescriptionInput: string;
};

export class TeamsBot extends TeamsActivityHandler {
  private readonly _dialog: AuthCommandDispatchDialog;

  constructor(
    private readonly _conversationState: ConversationState,
    private readonly _userState: UserState,
    readonly dedupStorage: Storage,
    private readonly _conversationReferences: SimpleConversationReferenceStore
  ) {
    super();

    this._dialog = new AuthCommandDispatchDialog(
      this._conversationState,
      dedupStorage
    );

    this.onMembersAdded(this._handleMembersAdded.bind(this));
    this.onInstallationUpdateAdd(this._handleInstalationUpdateAdd.bind(this));
    this.onInstallationUpdateRemove(
      this._handleInstalationUpdateRemove.bind(this)
    );
    this.onMessage(this._handleMessage.bind(this));
    this.onTokenResponseEvent(this._handleTokenResponse.bind(this));
  }

  /**
   * @inheritdoc
   */
  public async run(context: TurnContext) {
    await super.run(context).catch((error) => {
      console.error(
        `[${TeamsBot.name}][DEBUG] ${
          this.run.name
        } [[ERROR]]:\n${JSON.stringify(error, null, 2)}`
      );
    });

    // Save any state changes after the bot logic completes.
    await this._conversationState.saveChanges(context, false);
    await this._userState.saveChanges(context, false);
  }

  /**
   * @inheritdoc
   */
  public async handleTeamsSigninVerifyState(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this.handleTeamsSigninVerifyState.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this.handleTeamsSigninVerifyState.name
      } query:\n${JSON.stringify(query, null, 2)}`
    );

    await context.deleteActivity(context.activity.replyToId);

    const state = query.state;
    if (state.indexOf("CancelledByUser") >= 0) {
      await context.sendActivity("Consent flow was canceled by the user.");
      await this._dialog.end(context);
    } else {
      await this._dialog.run(context);
    }
  }

  /**
   * @inheritdoc
   */
  public async handleTeamsSigninTokenExchange(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this.handleTeamsSigninTokenExchange.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    await context.deleteActivity(context.activity.replyToId);
    await this._dialog.run(context);
  }

  /**
   * @inheritdoc
   */
  public async onInvokeActivity(context: TurnContext): Promise<any> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this.onInvokeActivity.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    // if (context.activity.name == tokenExchangeOperationName) {
    //   // const ssoToken = context.activity.value.token;
    // }

    if (context.activity.name == adaptiveCardActionName) {
      const activityValue: AdaptiveCardActionActivityValue =
        context.activity.value;

      console.debug(
        `[${TeamsBot.name}][DEBUG] ${
          this.onInvokeActivity.name
        } activityValue:\n${JSON.stringify(activityValue, null, 2)}`
      );

      if (activityValue.action.verb == "refresh") {
        await context.deleteActivity(context.activity.replyToId);
        await this._handleMessageInPersonalContext(
          context,
          activityValue.action.data?.command ?? "",
          activityValue.action.data
        );
      }

      if (activityValue.action.verb == "createTicket") {
        const actionData: AdaptiveCardActionActivityValueDataCreateTicket =
          activityValue.action.data;

        const cardJson = new ACData.Template(ticketCard).expand({
          $root: {
            ...actionData,
            showButtons: true,
            labelCancelButton: "Delete",
            enableCreateButton: false,
          },
        });
        const message = MessageFactory.attachment(
          CardFactory.adaptiveCard(cardJson)
        );
        message.id = context.activity.replyToId;
        await context.updateActivity(message);

        const graphClient = SimpleGraphClient.client(actionData.token);
        const threadMessages = await SimpleGraphClient.teamChannelMessages(
          graphClient,
          actionData.team.aadGroupId,
          actionData.channel.id,
          actionData.conversation.id
        );
        let oDataNextLink = threadMessages["@odata.nextLink"];
        while (oDataNextLink) {
          const nextThreadMessages =
            await SimpleGraphClient.teamChannelMessagesNext(
              graphClient,
              oDataNextLink
            );
          threadMessages.value.push(...nextThreadMessages.value);
          oDataNextLink = nextThreadMessages["@odata.nextLink"];
        }

        const result = {
          ticketStateChoiceSet: actionData.ticketStateChoiceSet,
          ticketCategoryChoiceSet: actionData.ticketCategoryChoiceSet,
          ticketDescriptionInput: actionData.ticketDescriptionInput,
          team: actionData.team,
          channel: actionData.channel,
          conversation: actionData.conversation,
          from: actionData.from,
          threadMessages: threadMessages.value.map(
            (message: TeamChannelMessage) => {
              return {
                attachments: message.attachments.map(
                  (attachment: TeamChannelMessageAttachment) => {
                    return {
                      contentType: attachment.contentType,
                      contentUrl: attachment.contentUrl,
                      content: attachment.content,
                      name: attachment.name,
                      teamsAppId: attachment.teamsAppId,
                      thumbnailUrl: attachment.thumbnailUrl,
                    } as TeamChannelMessageAttachment;
                  }
                ),
                from: {
                  displayName: message.from.displayName,
                  id: message.from.id,
                  tenantId: message.from.tenantId,
                  userIdentityType: message.from.userIdentityType,
                },
                id: message.id,
                createdDateTime: message.createdDateTime
                  ? new Date(message.createdDateTime)
                  : null,
                deletedDateTime: message.deletedDateTime
                  ? new Date(message.deletedDateTime)
                  : null,
                lastEditedDateTime: message.lastEditedDateTime
                  ? new Date(message.lastEditedDateTime)
                  : null,
                messageType: message.messageType,
                subject: message.subject,
                webUrl: message.webUrl,
              } as TeamChannelMessage;
            }
          ),
        };

        return await context.sendActivity(JSON.stringify(result, null, 2));
      }

      if (activityValue.action.verb == "cancelTicket") {
        return await context.deleteActivity(context.activity.replyToId);
      }
    }

    // Call super implementation for all other invoke activities.
    return await super.onInvokeActivity(context);
  }

  private async _handleInstalationUpdateAdd(
    context: TurnContext,
    next: () => Promise<void>
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleInstalationUpdateAdd.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    if (context.activity.conversation.conversationType !== "personal") {
      let pagedMembers: TeamsPagedMembersResult | null = null;
      let hasMore = true;

      while (hasMore) {
        pagedMembers = await TeamsInfo.getPagedTeamMembers(
          context,
          pagedMembers?.continuationToken
        );
        hasMore = !!pagedMembers.continuationToken;

        for (const member of pagedMembers.members) {
          console.debug(
            `[${TeamsBot.name}][DEBUG] ${
              this._handleInstalationUpdateAdd.name
            } member:\n${JSON.stringify(member, null, 2)}`
          );

          await AppInstallUtils.installAppInPersonalScope(
            context.activity.conversation.tenantId,
            member.aadObjectId
          );
        }
      }
    }

    // By calling next() you ensure that the next BotHandler is run.
    await next();
  }

  private async _handleInstalationUpdateRemove(
    context: TurnContext,
    next: () => Promise<void>
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleInstalationUpdateRemove.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    // By calling next() you ensure that the next BotHandler is run.
    await next();
  }

  private async _handleMembersAdded(
    context: TurnContext,
    next: () => Promise<void>
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleMembersAdded.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    const membersAdded = context.activity.membersAdded;
    for (const member of membersAdded) {
      if (member.id !== context.activity.recipient.id) {
        this._addConversationReference(context.activity);
      }
    }

    // By calling next() you ensure that the next BotHandler is run.
    await next();
  }

  private async _handleMessage(
    context: TurnContext,
    next: () => Promise<void>
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleMessage.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    // const userTokenClient = context.turnState.get(
    //   (<CloudAdapter>context.adapter).UserTokenClientKey
    // );
    // const magicCode =
    //   context.turnState && Number.isInteger(Number(context.turnState))
    //     ? context.turnState
    //     : "";
    // const tokenResponse = await userTokenClient.getUserToken(
    //   config.botConnectionName,
    //   context.activity.from.id,
    //   context.activity.channelId,
    //   magicCode
    // );

    let text = context.activity.text;
    // Remove the mention of this bot from activity text
    const removedMentionText = TurnContext.removeRecipientMention(
      context.activity
    );
    if (removedMentionText) {
      // Remove any line breaks
      text = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
    }

    console.debug(
      `[${TeamsBot.name}][DEBUG] ${this._handleMessage.name} text: ${text}`
    );
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${this._handleMessage.name} context.activity.conversation.conversationType: ${context.activity.conversation.conversationType}`
    );

    if (context.activity.conversation.conversationType !== "personal") {
      // Inside group chat, switch context to private chat with activity initiator
      await this._handleMessageGroupToPersonalContextSwitch(context, text);
    } else {
      // Inside personal chat, handle the incoming message
      await this._handleMessageInPersonalContext(context, text);
    }

    // By calling next() you ensure that the next BotHandler is run.
    await next();
  }

  private async _handleMessageGroupToPersonalContextSwitch(
    groupContext: TurnContext,
    command: string
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${this._handleMessageGroupToPersonalContextSwitch.name} switching to personal chat`
    );

    const teamDetails = await TeamsInfo.getTeamDetails(groupContext);
    const conversation = groupContext.activity.conversation;
    const channels = (await TeamsInfo.getTeamChannels(groupContext)).filter(
      (channel: ChannelInfo) =>
        channel.id ===
        (conversation?.id?.indexOf(";") >= 0
          ? conversation.id.split(";")[0]
          : conversation.id)
    );
    const channel = channels?.length > 0 ? channels[0] : { id: "", name: "" };
    const from = groupContext.activity.from;
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleMessageGroupToPersonalContextSwitch.name
      } teamDetails:\n${JSON.stringify(teamDetails, null, 2)}`
    );

    // Switch context to private chat with activity initiator
    const convoParams: ConversationParameters = {
      members: [groupContext.activity.from],
      isGroup: false,
      bot: groupContext.activity.recipient,
      tenantId: groupContext.activity.conversation.tenantId,
      activity: null,
      channelData: {
        tenant: { id: groupContext.activity.conversation.tenantId },
      },
    };
    await groupContext.adapter.createConversationAsync(
      config.botId,
      groupContext.activity.channelId,
      groupContext.activity.serviceUrl,
      null,
      convoParams,
      async (context: TurnContext): Promise<void> => {
        const conversationRef = TurnContext.getConversationReference(
          context.activity
        );

        console.debug(
          `[${TeamsBot.name}][DEBUG] ${
            this._handleMessageGroupToPersonalContextSwitch.name
          } createConversationAsync activity: \n${JSON.stringify(
            context.activity,
            null,
            2
          )}`
        );

        console.debug(
          `[${TeamsBot.name}][DEBUG] ${
            this._handleMessageGroupToPersonalContextSwitch.name
          } createConversationAsync conversationRef: \n${JSON.stringify(
            conversationRef,
            null,
            2
          )}`
        );

        await context.adapter.continueConversationAsync(
          config.botId,
          conversationRef,
          async (context: TurnContext) => {
            console.debug(
              `[${TeamsBot.name}][DEBUG] ${
                this._handleMessageGroupToPersonalContextSwitch.name
              } continueConversationAsync activity:\n${JSON.stringify(
                context.activity,
                null,
                2
              )}`
            );

            const cardJson = new ACData.Template(authRefreshCard).expand({
              $root: {
                team: teamDetails,
                channel: channel,
                conversation: conversation,
                from: from,
                userIds: [from.id],
              },
            });
            await context.sendActivity(
              MessageFactory.attachment(CardFactory.adaptiveCard(cardJson))
            );
          }
        );
      }
    );
  }

  private async _handleMessageInPersonalContext(
    context: TurnContext,
    command: string,
    data?: Partial<AdaptiveCardActionActivityValue>
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleMessageInPersonalContext.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    await this._dispatchAuthCommandInPersonalContext(context, {
      command,
      data,
    });
  }

  private async _dispatchAuthCommandInPersonalContext(
    context: TurnContext,
    data?: any
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._dispatchAuthCommandInPersonalContext.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    await this._dialog.run(context, data).catch((error) => {
      `[${TeamsBot.name}][DEBUG] ${
        this._dispatchAuthCommandInPersonalContext.name
      } [[ERROR]]:\n${JSON.stringify(error, null, 2)}`;
    });

    // const dialogContext = await this._dialogSet.createContext(context);
    // dialogContext.cancelAllDialogs();

    // let dialogTurnResult = await dialogContext.continueDialog();
    // if (dialogTurnResult?.status === DialogTurnStatus.empty) {
    //   dialogTurnResult = await dialogContext.beginDialog(INITIAL_DIALOG_ID);
    // }
    // return dialogTurnResult;
  }

  private async _handleTokenResponse(
    context: TurnContext,
    next: () => Promise<void>
  ): Promise<void> {
    console.debug(
      `[${TeamsBot.name}][DEBUG] ${
        this._handleTokenResponse.name
      } activity:\n${JSON.stringify(context.activity, null, 2)}`
    );

    await this._dialog.run(context);
    return await next();
  }

  private _addConversationReference(activity: Activity): void {
    const conversationReference =
      TurnContext.getConversationReference(activity);
    this._conversationReferences[conversationReference.user.aadObjectId] =
      conversationReference;
  }
}

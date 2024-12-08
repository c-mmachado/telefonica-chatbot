import {
  ComponentDialog,
  Dialog,
  DialogContext,
  DialogSet,
  DialogState,
  DialogTurnResult,
  DialogTurnStatus,
  OAuthPrompt,
  WaterfallDialog,
  WaterfallStepContext,
} from "botbuilder-dialogs";
import {
  ActivityTypes,
  ChannelAccount,
  ChannelInfo,
  ConversationAccount,
  ConversationState,
  MessageFactory,
  StatePropertyAccessor,
  Storage,
  TeamDetails,
  TurnContext,
  CardFactory,
  tokenExchangeOperationName,
  verifyStateOperationName,
  TokenResponse,
  InputHints,
} from "botbuilder";
import { ErrorCode, ErrorWithCode } from "@microsoft/teamsfx";
import * as ACData from "adaptivecards-templating";

import {
  SimpleGraphClient,
  TeamChannelMessage,
  TeamChannelResponse,
} from "../utils/graphClient";
import ticketCard from "../adaptiveCards/templates/createTicketCard.json";
import config from "../config/config";

export type WaterfallStepContextOptions<T> = {
  command: string;
  data?: T;
};

type AdaptiveCardActionSSORefreshData =
  | {
      team: TeamDetails;
      channel: ChannelInfo;
      conversation: ConversationAccount;
      from: ChannelAccount;
    }
  | undefined;

const MAIN_DIALOG = "MainDialog";
const INITIAL_DIALOG_ID = "MainWaterfallDialog";
// const TEAMS_BOT_SSO_PROMPT_ID = "TeamsBotSsoPrompt";
const OAUTH_PROMPT_ID = "OAuthPrompt";

const DIALOG_DATA = "dialogState";

export class AuthCommandDispatchDialog extends ComponentDialog {
  private readonly _dialogStateAccessor: StatePropertyAccessor<DialogState>;

  private readonly _dedupStorageKeys: string[];

  /**
   * Initializes a new instance of the MainDialog class.
   *
   * The dialog is composed of a waterfall dialog with the following steps:
   *  - PromptStep: Prompts the user to log in using SSO.
   *  - LoginStep: Verifies that the user has logged in.
   *
   *
   * The dialog makes use of TeamsBotSsoPrompt dialog that is used to handle the SSO login process.
   *
   * @param {ConversationState} conversationState The state of the conversation.
   */
  constructor(
    conversationState: ConversationState,
    private _dedupStorage: Storage
  ) {
    super(MAIN_DIALOG);

    this._dialogStateAccessor = conversationState.createProperty(DIALOG_DATA);

    // const settings: TeamsBotSsoPromptSettings = {
    //   scopes: [
    //     "User.Read",
    //     "Channel.ReadBasic.All",
    //     "Team.ReadBasic.All",
    //     "ChatMessage.Read",
    //     "ProfilePhoto.Read.All",
    //   ],
    //   timeout: 900000,
    //   endOnInvalidMessage: true,
    // };
    // const authConfig: OnBehalfOfCredentialAuthConfig = {
    //   authorityHost: config.authorityHost,
    //   clientId: config.clientId,
    //   tenantId: config.tenantId,
    //   clientSecret: config.clientSecret,
    // };
    // const loginUrl = `https://${config.botDomain}/auth-start.html`;
    // this.addDialog(
    //   new TeamsBotSsoPrompt(
    //     authConfig,
    //     loginUrl,
    //     TEAMS_BOT_SSO_PROMPT_ID,
    //     // {
    //     //   title: "Consent Flow",
    //     //   text: "Please review and accept the consent flow to continue.",
    //     //   timeout: 900000,
    //     //   endOnInvalidMessage: true,
    //     //   showSignInLink: true,
    //     //   connectionName: config.botConnectionName,
    //     // }
    //     settings
    //   )
    // );

    const oauthPrompt = new OAuthPrompt(
      OAUTH_PROMPT_ID,
      {
        title: "Consent Flow",
        text: "Please review and accept the consent flow to continue.",
        timeout: 900000,
        endOnInvalidMessage: true,
        showSignInLink: true,
        connectionName: config.botConnectionName,
      }
      // async (
      //   prompt: PromptValidatorContext<TokenResponse>
      // ): Promise<boolean> => {
      //   console.debug(
      //     `[${SSOCommandDispatchDialog.name}][DEBUG] [${
      //       OAuthPrompt.name
      //     }] promptValidator prompt:\n${JSON.stringify(prompt, null, 2)}`
      //   );
      //   return false;
      // }
    );
    oauthPrompt.beginDialog = this._cacheBypass;
    this.addDialog(oauthPrompt);

    this.addDialog(
      new WaterfallDialog<Partial<WaterfallStepContextOptions<any>>>(
        INITIAL_DIALOG_ID,
        [
          this._promptStep.bind(this),
          this._dedupStep.bind(this),
          this._dispatchStep.bind(this),
        ]
      )
    );

    this.initialDialogId = INITIAL_DIALOG_ID;
  }

  /**
   * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
   * If no dialog is active, it will start the default dialog.
   *
   * @param {TurnContext} context The context object for this turn of the conversation
   * @returns {Promise<DialogTurnResult>} A promise representing the result of the dialog's turn
   */
  public async run(
    context: TurnContext,
    data?: WaterfallStepContextOptions<any>
  ): Promise<DialogTurnResult> {
    // this._dialogStateAccessor.delete(context);
    const dialogSet = new DialogSet(this._dialogStateAccessor);
    dialogSet.add(this);
    const dialogContext = await dialogSet.createContext(context);
    // dialogContext.cancelAllDialogs();

    const dialogResult = await dialogContext.continueDialog();
    if (dialogResult?.status === DialogTurnStatus.empty) {
      return await dialogContext.beginDialog(this.id, data);
    }
    return dialogResult;
  }

  public async end(context: TurnContext): Promise<void> {
    const dialogSet = new DialogSet(this._dialogStateAccessor);
    dialogSet.add(this);
    const dialogContext = await dialogSet.createContext(context);
    // await this._dedupStorage.delete(this._dedupStorageKeys);
    await this._dialogStateAccessor.delete(context);
    await dialogContext.cancelAllDialogs();
  }

  private async _promptStep(
    stepContext: WaterfallStepContext<WaterfallStepContextOptions<any>>
  ): Promise<DialogTurnResult> {
    console.debug(
      `[${AuthCommandDispatchDialog.name}][DEBUG] [${WaterfallDialog.name}] ${this._promptStep.name}`
    );
    try {
      console.debug(
        `[${AuthCommandDispatchDialog.name}][DEBUG] [${WaterfallDialog.name}] ${
          this._promptStep.name
        } options:\n${JSON.stringify(stepContext.options, null, 2)}`
      );

      await stepContext
        .beginDialog(OAUTH_PROMPT_ID)
        .catch((error: any): Promise<DialogTurnResult> => {
          console.error(
            `[${AuthCommandDispatchDialog.name}][ERROR] [${
              WaterfallDialog.name
            }] ${this._promptStep.name} error:\n${JSON.stringify(
              error,
              null,
              2
            )}`
          );
          return stepContext.next();
        });
      return Dialog.EndOfTurn;
    } catch (error: any) {
      console.error(
        `[${AuthCommandDispatchDialog.name}][ERROR] [${WaterfallDialog.name}] ${
          this._promptStep.name
        } error:\n${JSON.stringify(error, null, 2)}`
      );
      return await stepContext.next();
    }
  }

  private async _dedupStep(
    stepContext: WaterfallStepContext<WaterfallStepContextOptions<any>>
  ): Promise<DialogTurnResult> {
    const tokenResult: Partial<TokenResponse> = stepContext.result;

    // Only dedup after promptStep to make sure that all Teams' clients receive the login request
    if (tokenResult && (await this._shouldDedup(stepContext.context))) {
      return Dialog.EndOfTurn;
    }
    return await stepContext.next(tokenResult);
  }

  private async _dispatchStep(
    stepContext: WaterfallStepContext<WaterfallStepContextOptions<any>>
  ): Promise<DialogTurnResult> {
    const tokenResult: Partial<TokenResponse> = stepContext.result;

    console.debug(
      `[${AuthCommandDispatchDialog.name}][DEBUG] [${WaterfallDialog.name}] ${this._dispatchStep.name}`
    );
    console.debug(
      `[${AuthCommandDispatchDialog.name}][DEBUG] [${WaterfallDialog.name}] ${
        this._dispatchStep.name
      } options:\n${JSON.stringify(stepContext.options, null, 2)}`
    );
    console.debug(
      `[${AuthCommandDispatchDialog.name}][DEBUG] [${WaterfallDialog.name}] ${
        this._dispatchStep.name
      } tokenResult:\n${JSON.stringify(tokenResult, null, 2)}`
    );

    if (tokenResult) {
      console.debug(
        `[${AuthCommandDispatchDialog.name}][DEBUG] [${WaterfallDialog.name}] ${
          this._dispatchStep.name
        } stepContext.context.activity:\n${JSON.stringify(
          stepContext.context.activity,
          null,
          2
        )}`
      );

      const command: string = stepContext.options?.command;
      // TODO: Handle command string and dispatch to appropriate command handler requiring SSO token.

      const data: Partial<AdaptiveCardActionSSORefreshData> =
        stepContext.options?.data;
      const graphClient = SimpleGraphClient.client(tokenResult.token);
      const userProfile = await SimpleGraphClient.me(graphClient);

      let channel: TeamChannelResponse | undefined;
      let message: TeamChannelMessage | undefined;
      let messageId: string | undefined;
      if (data?.team?.aadGroupId && data?.channel?.id) {
        channel = await SimpleGraphClient.teamChannel(
          graphClient,
          data.team.aadGroupId,
          data.channel.id
        );

        if (data.conversation?.id?.indexOf(";") >= 0) {
          messageId = data.conversation.id.split(";")[1];
          messageId = messageId.replace("messageid=", "");

          message = await SimpleGraphClient.teamChannelMessage(
            graphClient,
            data.team.aadGroupId,
            data.channel.id,
            messageId
          );
        }
      }

      // const userPhoto = await SimpleGraphClient.mePhoto(graphClient);

      const cardJson = new ACData.Template(ticketCard).expand({
        $root: {
          team: data?.team ?? { id: " ", name: " " },
          teamChoices: [],
          channel: {
            id: data?.channel?.id ?? " ",
            name: channel?.displayName ?? " ",
          },
          channelChoices: [],
          conversation: {
            id: messageId || " ",
            message: message?.subject || " ",
          },
          conversationChoices: [],
          from: data
            ? { ...data?.from, email: userProfile.mail }
            : { id: "", name: "", profileImage: "" },
          createdUtc: new Date().toUTCString(),
          token: tokenResult.token,
          showButtons: true,
          labelCancelButton: "Cancel",
          enableCreateButton: true
        },
      });

      await stepContext.context.sendActivity(
        MessageFactory.attachment(CardFactory.adaptiveCard(cardJson))
      );

      return await stepContext.endDialog(tokenResult);
    } else {
      await stepContext.context.sendActivity(
        `Unable to log you in or the authentication flow was rejected by the user.`
      );

      // Unable to retrieve token or an unexpected error occurred
      return await stepContext.endDialog();
    }
  }

  private _isSignInVerifyStateInvoke(context: TurnContext): boolean {
    const activity = context.activity;
    return (
      activity.type === ActivityTypes.Invoke &&
      activity.name === verifyStateOperationName
    );
  }

  private _isSignInTokenExchangeInvoke(context: TurnContext): boolean {
    const activity = context.activity;
    return (
      activity.type === ActivityTypes.Invoke &&
      activity.name === tokenExchangeOperationName
    );
  }

  // If a user is signed into multiple Teams clients, the Bot might receive a "signin/tokenExchange" from each client.
  // Each token exchange request for a specific user login will have an identical activity.value.Id.
  // Only one of these token exchange requests should be processed by the bot.  For a distributed bot in production,
  // this requires a distributed storage to ensure only one token exchange is processed.
  private async _shouldDedup(context: TurnContext): Promise<boolean> {
    if (
      (!this._isSignInTokenExchangeInvoke(context) &&
        !this._isSignInVerifyStateInvoke(context)) ||
      !context.activity.value?.id
    ) {
      return false;
    }

    const storeItem = {
      eTag: context.activity.value.id,
    };

    const key = this._getStorageKey(context);
    const storeItems = { [key]: storeItem };

    try {
      await this._dedupStorage.write(storeItems);
      this._dedupStorageKeys.push(key);
    } catch (error: any) {
      if (error instanceof Error && error.message.indexOf("eTag conflict")) {
        // Duplicate activity value id already in storage
        return true;
      }

      // Unexpected error encountered while writing to storage
      throw error;
    }
    return false;
  }

  private _getStorageKey(context: TurnContext): string {
    if (!context || !context.activity || !context.activity.conversation) {
      throw new Error("Unable to get storage key from current turn context");
    }
    const activity = context.activity;
    const channelId = activity.channelId;
    const conversationId = activity.conversation.id;

    if (
      !this._isSignInTokenExchangeInvoke(context) &&
      !this._isSignInVerifyStateInvoke(context)
    ) {
      throw new ErrorWithCode(
        `Unable to get storage key as current activity is of type 
        '${activity.type}::${activity.name}' and should be 
        '${ActivityTypes.Invoke}::${tokenExchangeOperationName} 
        or '${ActivityTypes.Invoke}::${verifyStateOperationName}'`,
        ErrorCode.FailedToRunDedupStep
      );
    }
    const value = activity.value;
    if (!value || !value.id) {
      throw new ErrorWithCode(
        "Unable to get storage key as current activity value is missing its id",
        ErrorCode.FailedToRunDedupStep
      );
    }
    return `${channelId}/${conversationId}/${value.id}`;
  }

  private async _cacheBypass(
    dc: DialogContext,
    options?: any
  ): Promise<DialogTurnResult> {
    // Ensure prompts have input hint set
    const o = Object.assign({}, options);
    if (
      o.prompt &&
      typeof o.prompt === "object" &&
      typeof o.prompt.inputHint !== "string"
    ) {
      o.prompt.inputHint = InputHints.AcceptingInput;
    }
    if (
      o.retryPrompt &&
      typeof o.retryPrompt === "object" &&
      typeof o.retryPrompt.inputHint !== "string"
    ) {
      o.retryPrompt.inputHint = InputHints.AcceptingInput;
    }
    // Initialize prompt state
    const timeout =
      typeof this["settings"].timeout === "number"
        ? this["settings"].timeout
        : 900000;
    const state = dc.activeDialog.state;
    state.state = {};
    state.options = o;
    state.expires = new Date().getTime() + timeout;
    // Attempt to get the users token
    // const output = yield UserTokenAccess.getUserToken(dc.context, this.settings, undefined);
    // if (output) {
    //     // Return token
    //     return yield dc.endDialog(output);
    // }
    // Prompt user to login
    await OAuthPrompt.sendOAuthCard(
      this["settings"],
      dc.context,
      state.options.prompt
    );
    return Dialog.EndOfTurn;
  }
}

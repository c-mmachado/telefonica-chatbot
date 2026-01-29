import {
    TeamsActivityHandler,
    TurnContext,
    ConversationState,
    UserState,
    SigninStateVerificationQuery,
    InvokeResponse,
    TeamsInfo,
    TeamsChannelAccount,
    MemoryStorage,
    // CloudAdapter,
} from "botbuilder";
// import { UserTokenClient } from "botframework-connector";

import { BotConfiguration } from "../config/config";
import { HandlerManager } from "./commands/manager";
import { DialogManager } from "./dialogs/manager";
import { AdaptiveCardAction, AdaptiveCardActionActivityValue } from "./adaptiveCards/actions/actions";
import { AdaptiveCards } from "./adaptiveCards/adaptiveCards";

import { TechnicianRepository } from "../server/repositories/technicians";
import { HandlerTurnContextFactory } from "./commands/context";

export interface TeamsBotOptions {
    config: BotConfiguration;

    conversationState: ConversationState;

    userState: UserState;

    handlerManager: HandlerManager;

    dialogManager: DialogManager;

    techRepository: TechnicianRepository;

    contextFactory: HandlerTurnContextFactory;
}

export interface TeamsBotBuilder {
    options(options: TeamsBotOptions): TeamsBotBuilder;

    config(config: BotConfiguration): TeamsBotBuilder;

    conversationState(conversationState: ConversationState): TeamsBotBuilder;

    userState(userState: UserState): TeamsBotBuilder;

    handlerManager(handlerManager: HandlerManager): TeamsBotBuilder;

    // dialogManager(dialogManager: DialogManager): TeamsBotBuilder;

    techRepository(techRepository: TechnicianRepository): TeamsBotBuilder;

    contextFactory(contextFactory: HandlerTurnContextFactory): TeamsBotBuilder;

    build(): TeamsBot;
}

export class DefaultTeamsBotBuilder implements TeamsBotBuilder {
    // public static Instance: TeamsBotBuilder = new TeamsBotBuilder();

    private _options: Partial<TeamsBotOptions> = {};

    constructor() {
        // if (TeamsBotBuilder.Instance) {
        //   throw new Error(
        //     `${TeamsBotBuilder.name} is a singleton class. Use '${TeamsBotBuilder.name}.Instance' to access the instance.`
        //   );
        // }
    }

    public options(options: TeamsBotOptions): TeamsBotBuilder {
        this._options = options;
        return this;
    }

    public config(config: BotConfiguration): TeamsBotBuilder {
        this._options.config = config;
        return this;
    }

    public conversationState(conversationState: ConversationState): TeamsBotBuilder {
        this._options.conversationState = conversationState;
        return this;
    }

    public userState(userState: UserState): TeamsBotBuilder {
        this._options.userState = userState;
        return this;
    }

    public handlerManager(handlerManager: HandlerManager): TeamsBotBuilder {
        this._options.handlerManager = handlerManager;
        return this;
    }

    // public dialogManager(dialogManager: DialogManager): TeamsBotBuilder {
    //     this._options.dialogManager = dialogManager;
    //     return this;
    // }

    public techRepository(techRepository: TechnicianRepository): TeamsBotBuilder {
        this._options.techRepository = techRepository;
        return this;
    }

    public contextFactory(contextFactory: HandlerTurnContextFactory): TeamsBotBuilder {
        this._options.contextFactory = contextFactory;
        return this;
    }

    public build(): TeamsBot {
        if (!this._options.config) {
            throw new Error(`Cannot build TeamsBot: missing 'config' option.`);
        }
        if (!this._options.conversationState) {
            this._options.conversationState = new ConversationState(new MemoryStorage());
        }
        if (!this._options.userState) {
            this._options.userState = new UserState(new MemoryStorage());
        }
        if (!this._options.handlerManager) {
            throw new Error(`Cannot build TeamsBot: missing 'handlerManager' option.`);
        }
        // if (!this._options.dialogManager) {
        //     throw new Error(`Cannot build TeamsBot: missing 'dialogManager' option.`);
        // }
        if (!this._options.techRepository) {
            throw new Error(`Cannot build TeamsBot: missing 'techRepository' option.`);
        }
        if (!this._options.contextFactory) {
            throw new Error(`Cannot build TeamsBot: missing 'contextFactory' option.`);
        }
        return new TeamsBot(
            this._options.config,
            this._options.conversationState,
            this._options.userState,
            this._options.handlerManager,
            // this._options.dialogManager,
            this._options.techRepository,
            this._options.contextFactory,
        );
    }
}

export class TeamsBot extends TeamsActivityHandler {
    constructor(
        private readonly _config: BotConfiguration,
        private readonly _conversationState: ConversationState,
        private readonly _userState: UserState,
        private readonly _handlerManager: HandlerManager,
        // private readonly _dialogManager: DialogManager,
        private readonly _techRepository: TechnicianRepository,
        private readonly _contextFactory: HandlerTurnContextFactory,
    ) {
        super();

        // this.onInstallationUpdateAdd(this._handleInstalationUpdateAdd.bind(this));
        // this.onInstallationUpdateRemove(
        //   this._handleInstalationUpdateRemove.bind(this)
        // );
        this.onMessage(this._handleMessage.bind(this));
        this.onMembersAdded(this._handleMembersAdded.bind(this));
        // this.onTokenResponseEvent(this._handleTokenResponse.bind(this));
    }

    public get config(): BotConfiguration {
        return this._config;
    }

    /**
     * @inheritdoc
     */
    public async run(context: TurnContext): Promise<void> {
        // Entry point for the bot logic which receives all incoming activities
        const proxiedContext: TurnContext = this._contextFactory.create(context);

        await super.run(proxiedContext).catch(async (error: any): Promise<void> => {
            await this._handleError(context, error).catch(async (err: any): Promise<void> => {
                console.error("Error handling run error:", err);
            });
        });

        // Save any state changes after the bot logic completes
        await this._conversationState.saveChanges(proxiedContext, false);
        await this._userState.saveChanges(proxiedContext, false);
    }

    /**
     * @inheritdoc
     */
    public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse<any>> {
        console.debug("context.activity:", context.activity);

        if (context.activity.name === AdaptiveCardAction.Name) {
            // Extracts the action value from the activity when the activity has name 'adaptiveCard/action'
            const value: AdaptiveCardActionActivityValue = context.activity.value;

            // Resolves action handler from 'activity.value.action.verb' and dispatches the action
            const cardOrText: any | string = await this._handlerManager.resolveAndDispatch<any | string>(
                context,
                value.action.verb,
                value.action.data,
            );

            // Return an invoke response to indicate that the activity was handled and to prevent the Teams client from displaying an error message
            // due to the activity not being responded to
            // return { status: StatusCodes.OK };
            // TODO: Check if cardOrText is string or Adaptive Card and return appropriate InvokeResponse object as well as if there is an Error instead
            return AdaptiveCards.invokeResponse(cardOrText);
        }

        // Call super implementation for all other invoke activities
        return await super.onInvokeActivity(context);
    }

    /**
     * @inheritdoc
     */
    public async handleTeamsSigninVerifyState(
        context: TurnContext,
        query: SigninStateVerificationQuery,
    ): Promise<void> {
        return await this._onSignInAction(context, query);
    }

    /**
     * @inheritdoc
     */
    public async handleTeamsSigninTokenExchange(
        context: TurnContext,
        query: SigninStateVerificationQuery,
    ): Promise<void> {
        return await this._onSignInAction(context, query);
    }

    private async _onSignInAction(context: TurnContext, query: SigninStateVerificationQuery): Promise<void> {
        return this._handlerManager.onSignInAction(context, query).catch(async (error: any): Promise<void> => {
            await this._handleError(context, error).catch(async (err: any): Promise<void> => {
                console.error("Error handling sign-in action error:", err);
            });
        });
    }

    // private async _handleTokenResponse(context: TurnContext, next: () => Promise<void>): Promise<void> {
    //     // This activity type can be triggered during an SSO flow (Currently unused)
    //     console.debug(`context.activity:`, context.activity);

    //     if (context.activity?.replyToId) {
    //         await context.deleteActivity(context.activity.replyToId);
    //     }

    //     await this._dialogManager.continueDialog(context, OAuthDialog.name).catch((error: any): void => {
    //         console.error(error);
    //         while (error?.cause) {
    //             error = error.cause;
    //             console.error("Caused by:", error);
    //         }
    //     });

    //     return await next();
    // }

    private async _handleMessage(context: TurnContext, next: () => Promise<void>): Promise<void> {
        console.debug("context.activity:", context.activity);

        // Removes the mention of this bot from activity text
        let text = context.activity.text;
        const removedMentionText = TurnContext.removeRecipientMention(context.activity);
        if (removedMentionText) {
            text = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
        }

        if (!text || text.length === 0) {
            // If the text is empty, check if activity value is present and contains an 'action.verb'
            console.warn(`Empty message text`);

            if (context.activity.value?.action) {
                // If the activity value contains an action, delegate the handling to the 'onInvokeActivity' method
                // Set the activity name to 'adaptiveCard/action' to trigger the onInvokeActivity method
                context.activity.name = AdaptiveCardAction.Name;
                await this.onInvokeActivity(context);

                return await next();
            }
        }

        console.debug(`text: '${text}'`);
        console.debug(
            `context.activity.conversation.conversationType: '${context.activity.conversation.conversationType}'`,
        );

        // Gets the caller information
        const fromInfo: TeamsChannelAccount = await TeamsInfo.getMember(context, context.activity.from.id);
        if (!fromInfo) {
            // If the caller email address cannot be resolved, log an error and return as it would be impossible to validate
            // if the caller is a technician
            console.error(
                `Unable to resolve caller email address for user '${context.activity.from.name}' with id '${context.activity.from.id}' and aadObjectId '${context.activity.from.aadObjectId}'`,
            );
            throw new Error(
                `Unable to resolve caller email address for user '${context.activity.from.name}' with id '${context.activity.from.id}' and aadObjectId '${context.activity.from.aadObjectId}'`,
            );
        }

        if (!this.config.allowAll) {
            // If the bot is not configured to allow all users, check if the caller is a technician
            const technician = fromInfo.email ? await this._techRepository.technicianByEmail(fromInfo.email) : null;
            if (!technician) {
                // If the caller is not a technician, log a warning and return as the caller is not authorized to use the bot
                console.warn(`Caller '${fromInfo.email}' is not registered as a technician`);

                return await next();
            }
        }

        await this._handlerManager.resolveAndDispatch(context, text).catch(async (error: any): Promise<void> => {
            await this._handleError(context, error).catch(async (err: any): Promise<void> => {
                console.error("Error handling message dispatch error:", err);
            });
        });

        return await next();
    }

    private async _handleMembersAdded(context: TurnContext, next: () => Promise<void>): Promise<void> {
        console.debug("context.activity:", context.activity);

        const membersAdded = context.activity.membersAdded;
        for (const member of membersAdded ?? []) {
            // Greet anyone that was not the target (recipient) of this message
            if (member.id !== context.activity.recipient.id) {
                // const welcomeText = `¡Hola y bienvenido! Soy el bot de gestión de tickets. ¿En qué puedo ayudarte hoy?`;
                // await context.sendActivity(welcomeText);
            }
        }
        return await next();
    }

    private async _handleError(context: TurnContext, error: any): Promise<void> {
        let errorMsg = `Hay ocurrido un error al procesar la actividad. Por favor, inténtalo de nuevo más tarde.\n\n Razón: ${error.message}\n\n`;

        console.error(error);

        while (error?.cause || error?.reason) {
            error = error.cause || error.reason;
            errorMsg += `Causado por: '${error.message}'\n`;

            console.error("Caused by:", error);
        }

        if (errorMsg.length > 0) {
            await context.sendActivity(errorMsg);
        }
    }
}

}

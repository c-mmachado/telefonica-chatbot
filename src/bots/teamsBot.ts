import {
    TeamsActivityHandler,
    TurnContext,
    ConversationState,
    UserState,
    SigninStateVerificationQuery,
    InvokeResponse,
    TeamsInfo,
    TeamsChannelAccount,
    // MemoryStorage,
    // CloudAdapter,
} from "botbuilder";
// import { UserTokenClient } from "botframework-connector";

import { BotConfiguration } from "../config/config";
import { HandlerManager } from "./commands/manager";
// import { DialogManager } from "./dialogs/manager";
import { AdaptiveCardAction, AdaptiveCardActionActivityValue } from "./adaptiveCards/actions/actions";
import { AdaptiveCards } from "./adaptiveCards/adaptiveCards";

import { TechnicianRepository } from "../server/repositories/technicians";
import { HandlerTurnContextFactory } from "./commands/context";

export abstract class ActivityHandlerFactory<ParamTypes extends Array<any> = []> {
    public static Default: ActivityHandlerFactory<[DefaultActivityHandlerOptions]>;

    public abstract create(...options: ParamTypes): TeamsActivityHandler;
}

interface DefaultActivityHandlerOptions {
    config: BotConfiguration;

    conversationState: ConversationState;

    userState: UserState;

    handlerManager: HandlerManager;

    // dialogManager: DialogManager;

    techRepository: TechnicianRepository;

    contextFactory: HandlerTurnContextFactory;

    errorHandler: ActivityErrorHandler;
}

// @ts-ignore
class DefaultActivityHandlerFactory extends ActivityHandlerFactory<[DefaultActivityHandlerOptions]> {
    private static _instance: DefaultActivityHandlerFactory = new DefaultActivityHandlerFactory();

    protected constructor() {
        super();
        ActivityHandlerFactory.Default = DefaultActivityHandlerFactory._instance;
    }

    public create(options: DefaultActivityHandlerOptions): TeamsActivityHandler {
        return new TeamsBot(
            options.config,
            options.conversationState,
            options.userState,
            options.handlerManager,
            // options.dialogManager,
            options.techRepository,
            options.contextFactory,
            options.errorHandler,
        );
    }
}

export abstract class ActivityErrorHandlerFactory<ParamTypes extends Array<any> = []> {
    public static Default: ActivityErrorHandlerFactory<[]>;

    abstract create(...options: ParamTypes): ActivityErrorHandler;
}

// @ts-ignore
class DefaultActivityErrorHandlerFactory extends ActivityErrorHandlerFactory<[]> {
    private static _instance: DefaultActivityErrorHandlerFactory = new DefaultActivityErrorHandlerFactory();

    protected constructor() {
        super();
        ActivityErrorHandlerFactory.Default = DefaultActivityErrorHandlerFactory._instance;
    }

    public create(): ActivityErrorHandler {
        return new DefaultActivityErrorHandler();
    }
}

export interface ActivityErrorHandler {
    handle(context: TurnContext, error: any): Promise<void>;
}

class DefaultActivityErrorHandler implements ActivityErrorHandler {
    public async handle(context: TurnContext, error: any): Promise<void> {
        try {
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
        } catch (error: any) {
            console.error("Error handling activity error:", error);
            // Swallow the error to prevent bot from crashing
        }
    }
}

class TeamsBot extends TeamsActivityHandler {
    constructor(
        private readonly _config: BotConfiguration,
        private readonly _conversationState: ConversationState,
        private readonly _userState: UserState,
        private readonly _handlerManager: HandlerManager,
        // private readonly _dialogManager: DialogManager,
        private readonly _techRepository: TechnicianRepository,
        private readonly _contextFactory: HandlerTurnContextFactory,
        private readonly _errorHandler: ActivityErrorHandler,
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
            await this._errorHandler.handle(context, error).catch(async (err: any): Promise<void> => {
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
            await this._errorHandler.handle(context, error).catch(async (err: any): Promise<void> => {
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
            await this._errorHandler.handle(context, error).catch(async (err: any): Promise<void> => {
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

    // private async _handleError(context: TurnContext, error: any): Promise<void> {
    //     let errorMsg = `Hay ocurrido un error al procesar la actividad. Por favor, inténtalo de nuevo más tarde.\n\n Razón: ${error.message}\n\n`;

    //     console.error(error);

    //     while (error?.cause || error?.reason) {
    //         error = error.cause || error.reason;
    //         errorMsg += `Causado por: '${error.message}'\n`;

    //         console.error("Caused by:", error);
    //     }

    //     if (errorMsg.length > 0) {
    //         await context.sendActivity(errorMsg);
    //     }
    // }
}

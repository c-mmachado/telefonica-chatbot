import { ChannelAccount, ConversationParameters, ConversationReference, TurnContext } from "botbuilder";

import { BotConfiguration } from "../../config/config";

export interface ConversationReferenceStore {
    [key: string]: Partial<ConversationReference>;
}

export interface ConversationHelper {
    switchToConversation(
        context: TurnContext,
        conversationReference: Partial<ConversationReference>,
        action: (context: TurnContext) => Promise<any>
    ): Promise<void>;

    switchToPersonalConversation(context: TurnContext, action: (context: TurnContext) => Promise<any>): Promise<void>;

    reference(reference: Partial<ConversationReference>): Partial<ConversationReference>;
    reference(user: ChannelAccount): Partial<ConversationReference>;
}

export class DefaultConversationHelper implements ConversationHelper {
    constructor(
        private readonly _config: BotConfiguration,
        private readonly _conversationStore: ConversationReferenceStore
    ) {}

    public async switchToConversation(
        context: TurnContext,
        conversationReference: Partial<ConversationReference>, //| ConversationParameters
        action: (context: TurnContext) => Promise<any>
    ): Promise<void> {
        await context.adapter.continueConversationAsync(
            this._config.botId,
            conversationReference,
            async (chatOrGroupContext: TurnContext): Promise<void> => {
                await action?.(chatOrGroupContext);
            }
        );
    }

    public async switchToPersonalConversation(
        context: TurnContext,
        action: (context: TurnContext) => Promise<any>
    ): Promise<void> {
        if (context.activity.conversation.conversationType === "personal") {
            return await action?.(context);
        }

        // Get the conversation reference for the user
        const reference = this.reference(context.activity.from);
        if (reference) {
            // If the conversation reference is available, switch context to the private chat
            // using the conversation reference
            await context.adapter.continueConversationAsync(
                this._config.botId,
                reference,
                async (chatContext: TurnContext) => {
                    console.debug(`chatContext.activity:`, chatContext.activity);
                    await action?.(chatContext);
                }
            );
        } else {
            // If the conversation reference is not available, create a new conversation
            // and switch context to the private chat with the activity initiator
            const convoParams: ConversationParameters = {
                members: [context.activity.from],
                isGroup: false,
                bot: context.activity.recipient,
                tenantId: context.activity.conversation.tenantId,
                channelData: {
                    tenant: { id: context.activity.conversation.tenantId },
                },
            };

            // Create a new conversation with the user and bot in a private chat
            await context.adapter.createConversationAsync(
                this._config.botId,
                context.activity.channelId,
                context.activity.serviceUrl,
                null as any,
                convoParams,
                async (createContext: TurnContext): Promise<void> => {
                    // Gets the newly created conversation reference for the user and store it
                    const reference = TurnContext.getConversationReference(createContext.activity);
                    this.reference(reference);

                    console.debug(`activity:`, createContext.activity);
                    // await action?.(createContext);

                    // Continue the conversation in the private chat
                    await createContext.adapter.continueConversationAsync(
                        this._config.botId,
                        reference,
                        async (chatContext: TurnContext) => {
                            console.debug(`activity:`, chatContext.activity);
                            await action?.(chatContext);
                        }
                    );
                }
            );
        }
    }

    public reference(reference: Partial<ConversationReference>): Partial<ConversationReference>;
    public reference(user: ChannelAccount): Partial<ConversationReference>;
    public reference(
        referenceOrUser: Partial<ConversationReference> | ChannelAccount
    ): Partial<ConversationReference> | null {
        if ("user" in referenceOrUser && referenceOrUser.user?.aadObjectId) {
            // Store the conversation reference in memory using the user id as key
            // if the user id is available in the conversation
            this._conversationStore[referenceOrUser.user.aadObjectId] = referenceOrUser;
            return referenceOrUser;
        }

        if ("aadObjectId" in referenceOrUser && referenceOrUser.aadObjectId) {
            return this._conversationStore[referenceOrUser.aadObjectId];
        }

        return null;
    }
}

import { ConversationReference, TurnContext } from "botbuilder";
import { DialogTurnResult } from "botbuilder-dialogs";

import {
  ActiveRequest,
  DefaultHandlerManager,
  HandlerTriggerData,
} from "./manager";
import { ConversationHelper } from "./conversation";
import { DialogManager } from "../dialogs/manager";

export interface HandlerTurnContextFactory {
  create(context: TurnContext): TurnContext;
}

export class DefaultHandlerTurnContextFactory
  implements HandlerTurnContextFactory
{
  constructor(
    private readonly _dialogManager: DialogManager,
    private readonly _conversationHelper: ConversationHelper
  ) {}

  public create(context: TurnContext): TurnContext {
    return new DefaultHandlerTurnContextBuilder()
      .context(context)
      .dialogManager(this._dialogManager)
      .contextHelper(this._conversationHelper)
      .contextFactory(this)
      .build();
  }
}

export interface HandlerTurnContextBuilder {
  context(context: TurnContext): HandlerTurnContextBuilder;

  dialogManager(dialogManager: DialogManager): HandlerTurnContextBuilder;

  contextHelper(contextHelper: ConversationHelper): HandlerTurnContextBuilder;

  contextFactory(
    contextFactory: HandlerTurnContextFactory
  ): HandlerTurnContextBuilder;

  build(): TurnContext;
}

export class DefaultHandlerTurnContextBuilder
  implements HandlerTurnContextBuilder
{
  private _options: Partial<DefaultHandlerTurnContextOptions> = {};

  public context(context: TurnContext): HandlerTurnContextBuilder {
    this._options.context = context;
    return this;
  }

  public dialogManager(
    dialogManager: DialogManager
  ): HandlerTurnContextBuilder {
    this._options.dialogManager = dialogManager;
    return this;
  }

  public contextHelper(
    contextHelper: ConversationHelper
  ): HandlerTurnContextBuilder {
    this._options.contextHelper = contextHelper;
    return this;
  }

  public contextFactory(
    contextFactory: HandlerTurnContextFactory
  ): HandlerTurnContextBuilder {
    this._options.contextFactory = contextFactory;
    return this;
  }

  public build(): TurnContext {
    if (!this._options.context) {
      throw new Error(`Argument 'context' cannot be null or undefined`);
    }

    if (!this._options.dialogManager) {
      throw new Error(`Argument 'dialogManager' cannot be null or undefined`);
    }

    if (!this._options.contextHelper) {
      throw new Error(`Argument 'contextHelper' cannot be null or undefined`);
    }

    if (!this._options.contextFactory) {
      throw new Error(`Argument 'contextFactory' cannot be null or undefined`);
    }

    return DefaultHandlerTurnContext.from(
      this._options.context,
      this._options as DefaultHandlerTurnContextOptions
    );
  }
}

export interface HandlerTurnContext {
  switchToConversation(
    conversationReference: Partial<ConversationReference>,
    action: (chatOrGroupContext: TurnContext) => Promise<void>
  ): Promise<void>;

  switchToPersonalConversation(
    action: (chatContext: TurnContext) => Promise<void>
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

  trigger(): Partial<HandlerTriggerData>;

  request<T>(): ActiveRequest<T>;
}

export interface DefaultHandlerTurnContextOptions {
  context: TurnContext;

  dialogManager: DialogManager;

  contextHelper: ConversationHelper;

  contextFactory: HandlerTurnContextFactory;
}

export class DefaultHandlerTurnContext implements HandlerTurnContext {
  public static from(
    context: TurnContext,
    options: DefaultHandlerTurnContextOptions
  ): HandlerTurnContext & TurnContext {
    const extended: HandlerTurnContext = new DefaultHandlerTurnContext(
      options.context,
      options.dialogManager,
      options.contextHelper,
      options.contextFactory
    );

    // Proxies 'HandlerTurnContext' and 'TurnContext' into a single object instance to
    // allow access to both sets of properties and methods transparently.
    const proxy = new Proxy(context, {
      get(target: TurnContext, prop: string | symbol, receiver: any): any {
        // Try 'HandlerTurnContext' instance
        if (prop in extended) {
          const value = (extended as any)[prop];
          if (typeof value === "function") {
            return value.bind(extended);
          }
          return value;
        }

        // Fallback to the real 'TurnContext'
        const value = Reflect.get(target, prop, receiver);
        if (typeof value === "function") {
          return value.bind(target);
        }

        return value;
      },

      set(target: TurnContext, prop: string | symbol, value: any): boolean {
        if (prop in extended) {
          (extended as any)[prop] = value;
          return true;
        }
        // Let the original TurnContext handle its own properties
        return Reflect.set(target, prop, value);
      },
    });

    return proxy as unknown as HandlerTurnContext & TurnContext;
  }

  private constructor(
    private readonly _context: TurnContext,
    private readonly _dialogManager: DialogManager,
    private readonly _contextHelper: ConversationHelper,
    private readonly _contextFactory: HandlerTurnContextFactory
  ) {}

  public trigger(): Partial<HandlerTriggerData> {
    return this.request()?.trigger ?? {};
  }

  public request<T>(): ActiveRequest<T> {
    return this._context.turnState.get(DefaultHandlerManager.ActiveRequestKey);
  }

  public async switchToConversation(
    reference: Partial<ConversationReference>,
    action: (context: TurnContext) => Promise<void>
  ): Promise<void> {
    // Switch to the specified conversation and execute the action
    await this._contextHelper.switchToConversation(
      this._context,
      reference,
      async (chatOrGroupContext: TurnContext): Promise<void> => {
        await action?.(this._contextFactory.create(chatOrGroupContext));
      }
    );
  }

  public async switchToPersonalConversation(
    action: (context: TurnContext) => Promise<void>
  ): Promise<void> {
    if (this._context.activity.conversation.conversationType === "personal") {
      return await action?.(this._contextFactory.create(this._context));
    }

    // Switch to personal context and execute action
    await this._contextHelper.switchToPersonalConversation(
      this._context,
      async (chatContext: TurnContext): Promise<void> => {
        return await action?.(this._contextFactory.create(chatContext));
      }
    );
  }

  public async runDialog(
    context: TurnContext,
    dialogName: string,
    data?: any
  ): Promise<DialogTurnResult> {
    return await this._dialogManager.runDialog(context, dialogName, data);
  }

  public async continueDialog(
    context: TurnContext,
    dialogName: string
  ): Promise<DialogTurnResult> {
    return await this._dialogManager.continueDialog(context, dialogName);
  }

  public async stopDialog(
    context: TurnContext,
    dialogName: string
  ): Promise<DialogTurnResult> {
    return await this._dialogManager.stopDialog(context, dialogName);
  }
}

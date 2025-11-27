import {
  ActivityTypes,
  AdaptiveCardInvokeResponse,
  CardFactory,
  MessageFactory,
  StatusCodes,
  TurnContext,
} from "botbuilder";
import {
  MicrosoftAppCredentials,
  ConnectorClient,
} from "botframework-connector";

import { config } from "../../config/config";

export class AdaptiveCards {
  public static readonly Type: string =
    "application/vnd.microsoft.card.adaptive";

  constructor() {
    throw new Error(
      `${AdaptiveCards.name} cannot be instantiated as it is a utility class.`
    );
  }

  public static invokeResponse(
    card: any
  ): Partial<AdaptiveCardInvokeResponse> & { status: number } {
    const cardResponse = {
      type: AdaptiveCards.Type,
      value: card,
    };

    return {
      statusCode: StatusCodes.OK,
      status: StatusCodes.OK,
      type: ActivityTypes.InvokeResponse,
      body: cardResponse,
    } as Partial<AdaptiveCardInvokeResponse> & { status: number };
  }

  public static async updateCard(
    context: TurnContext,
    card: any
  ): Promise<void> {
    const serviceUrl = context.activity.serviceUrl;
    const credentials = new MicrosoftAppCredentials(
      config.botId,
      config.botPassword
    );
    const connectorClient = new ConnectorClient(credentials, {
      baseUri: serviceUrl,
    });

    const conversationId = context.activity.conversation.id;
    const activityId = context.activity.replyToId;
    if (!activityId) {
      throw new Error(
        `[${AdaptiveCards.name}][ERROR] ${this.updateCard.name} 'context.activity.replyToId' is undefined`
      );
    }

    const replyActivity = MessageFactory.attachment(
      CardFactory.adaptiveCard(card)
    );
    replyActivity.id = activityId;
    replyActivity.attachments = [CardFactory.adaptiveCard(card)];

    await connectorClient.conversations.updateActivity(
      conversationId,
      activityId,
      replyActivity
    );
  }
}

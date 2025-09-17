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
import { config } from "../config/config";

export class AdaptiveCards {
  constructor() {
    throw new Error(
      `${AdaptiveCards.name} cannot be instantiated as it is a utility class.`
    );
  }

  public static invokeResponse(
    card: any
  ): Partial<AdaptiveCardInvokeResponse> & { status: number } {
    const cardResponse = {
      type: "application/vnd.microsoft.card.adaptive",
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

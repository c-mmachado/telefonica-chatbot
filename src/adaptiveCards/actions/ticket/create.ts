import {
  CardFactory,
  ConversationAccount,
  MessageFactory,
  TeamDetails,
} from "botbuilder";
import { TriggerPatterns, CommandMessage } from "@microsoft/teamsfx";

import * as ACData from "adaptivecards-templating";

import { ActionHandler, HandlerTurnContext } from "../../../commands/handler";
import {
  ApplicationIdentityType,
  SimpleGraphClient,
  TeamsChannelMessage,
} from "../../../utils/graphClient";
import { APIClient, Queue, Ticket } from "../../../utils/apiClient";
import { BotConfiguration } from "../../../config/config";

import ticketCard from "../../../adaptiveCards/templates/ticketCard.json";

type AdaptiveCardActionCreateTicketData = {
  command: string;
  team: TeamDetails & { choices: { title: string; value: string }[] };
  channel: { id: string; name: string } & {
    choices: { title: string; value: string }[];
  };
  conversation: { id: string; name: string } & {
    choices: { title: string; value: string }[];
  };
  from: ConversationAccount & { choices: { title: string; value: string }[] };
  ticket: {
    state: {
      id: string;
      choices: { title: string; value: string }[];
    };
    queue: {
      id: string;
      choices: { title: string; value: string }[];
    };
    description: string;
  };
  token: string;
  createdUtc: string;
  gui: any;

  ticketStateChoiceSet: string;
  ticketCategoryChoiceSet: string;
  ticketDescriptionInput: string;
};

export class TicketAdaptiveCardCreateActionHandler implements ActionHandler {
  public pattern: TriggerPatterns = "createTicket";

  constructor(
    private readonly _config: BotConfiguration,
    private readonly _apiClient: APIClient
  ) {}

  public async run(
    handlerContext: HandlerTurnContext,
    commandMessage: CommandMessage,
    data?: any
  ): Promise<any> {
    console.debug(
      `[${TicketAdaptiveCardCreateActionHandler.name}][DEBUG] [${this.run.name}]`
    );

    const actionData: AdaptiveCardActionCreateTicketData =
      handlerContext.context.activity.value?.action?.data;

    const cardJson = new ACData.Template(ticketCard).expand({
      $root: {
        ...actionData,
        ticket: {
          state: {
            id: actionData.ticketStateChoiceSet,
            choices: actionData.ticket.state.choices.filter(
              (v) => v.value == actionData.ticketStateChoiceSet
            ),
          },
          queue: {
            id: actionData.ticketCategoryChoiceSet,
            choices: actionData.ticket.queue.choices.filter(
              (v) => v.value == actionData.ticketCategoryChoiceSet
            ),
          },
          description: actionData.ticketDescriptionInput,
        },
        gui: {
          buttons: {
            visible: true,
            create: {
              ...actionData.gui.buttons.create,
              enabled: false,
            },
            cancel: {
              ...actionData.gui.buttons.cancel,
              label: "Borrar",
            },
          },
        },
      },
    });
    const message = MessageFactory.attachment(
      CardFactory.adaptiveCard(cardJson)
    );
    message.id = handlerContext.context.activity.replyToId;
    await handlerContext.context.updateActivity(message);

    const graphClient = SimpleGraphClient.client(actionData.token);
    const initialMessage = await SimpleGraphClient.teamsChannelMessage(
      graphClient,
      actionData.team.aadGroupId,
      actionData.channel.id,
      actionData.conversation.id
    );
    const threadMessages = await SimpleGraphClient.teamsChannelMessages(
      graphClient,
      actionData.team.aadGroupId,
      actionData.channel.id,
      actionData.conversation.id
    );
    let oDataNextLink = threadMessages["@odata.nextLink"];
    while (oDataNextLink) {
      const nextThreadMessages =
        await SimpleGraphClient.teamsChannelMessagesNext(
          graphClient,
          oDataNextLink
        );
      threadMessages.value.push(...nextThreadMessages.value);
      oDataNextLink = nextThreadMessages["@odata.nextLink"];
    }
    threadMessages.value = [
      {
        body: {
          content: actionData.ticketDescriptionInput,
          contentType: "text/html",
        },
        from: initialMessage.from,
      } as TeamsChannelMessage,
      initialMessage,
      ...threadMessages.value.reverse(),
    ];

    console.debug(
      `[${TicketAdaptiveCardCreateActionHandler.name}][DEBUG] ${this.run.name} threadMessages.length: ${threadMessages.value.length}`
    );

    // TeamsInfo.getPagedTeamMembers(context, undefined).then((result) => {});
    const queue: Queue = await this._apiClient.queue(
      actionData.ticketCategoryChoiceSet
    );
    const ticket = await this._apiClient.createTicket(
      queue,
      initialMessage.subject
    );
    // const ticket: Ticket = await this._apiClient.ticket({
    //   id: "416115",
    //   _url: "https://test-epg-vmticket-01.hi.inet/REST/2.0/ticket/416115",
    //   type: "ticket",
    // });
    // const ticket: Partial<Ticket> = {
    //   id: "416115",
    //   _hyperlinks: [
    //     {
    //       ref: "comment",
    //       _url: "https://test-epg-vmticket-01.hi.inet/REST/2.0/ticket/416115/comment",
    //       type: "comment",
    //     },
    //   ],
    // };

    console.debug(
      `[${TicketAdaptiveCardCreateActionHandler.name}][DEBUG] ${
        this.run.name
      } ticket:\n${JSON.stringify(ticket, null, 2)}`
    );

    for (const message of threadMessages.value) {
      if (!message.body?.content?.trim() || !message.from?.user) {
        continue;
      }

      console.debug(
        `[${TicketAdaptiveCardCreateActionHandler.name}][DEBUG] ${
          this.run.name
        } message:\n${JSON.stringify(message, null, 2)}`
      );

      if (message.mentions?.length === 1) {
        if (
          message.mentions[0].mentioned?.application
            ?.applicationIdentityType === ApplicationIdentityType.BOT &&
          message.mentions[0].mentioned?.application?.id === this._config.botId
        ) {
          console.debug(
            `[${TicketAdaptiveCardCreateActionHandler.name}][DEBUG] ${this.run.name} message mentions bot, skipping...`
          );
          continue;
        }
      }

      await this._apiClient.addTicketComment(graphClient, actionData.token, ticket, message);
    }

    return await handlerContext.context.sendActivity(
      `Se hay creado el ticket con el número: ${ticket.id}. Lo puedes acceder en [este enlace](https://test-epg-vmticket-01.hi.inet/Ticket/Display.html?id=${ticket.id}).`
    );
  }
}

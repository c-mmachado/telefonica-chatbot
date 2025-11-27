import { CardFactory, MessageFactory, TurnContext } from "botbuilder";
import { ChatMessage } from "@microsoft/microsoft-graph-types";

import * as ACData from "adaptivecards-templating";

import {
    AdaptiveCardActionActivityValue,
    AdaptiveCardActionPositiveTicketPageData,
    AdaptiveCardTicketCardPageData,
} from "../actions";
import { ActionHandler } from "../../../commands/handler";
import { HandlerMessage } from "../../../commands/message";
import { HandlerTriggerData } from "../../../commands/manager";
import { BotConfiguration } from "../../../../config/config";
// import { RTClient } from "../../../../utils/client/rt/rt";
// import { Queue, Ticket } from "../../../../utils/client/rt/model";
import { GraphClient } from "../../../../utils/client/graph";
import { isKeyOf } from "../../../../utils/misc";
// import { LogsRepository } from "../../../server/repositories/logs";

import page1 from "../../templates/ticket/page1.json";

export class TicketAdaptiveCardCreateActionHandler implements ActionHandler {
    public pattern: string = "createTicket";

    constructor(
        private readonly _config: BotConfiguration,
        // private readonly _rt: RTClient,
        private readonly _graph: GraphClient // private readonly _logs: LogsRepository
    ) {}

    public async run(context: TurnContext, _handlerMessage: HandlerMessage): Promise<any> {
        const activityValue: AdaptiveCardActionActivityValue = context.activity.value;
        const actionData: AdaptiveCardActionPositiveTicketPageData = activityValue?.action?.data;
        const state: Record<string, any> = (context as any).request().data;

        if (state.gui.page === 1) {
            for (const [key, value] of Object.entries<any>(state.ticket.customFields)) {
                if (isKeyOf(key, actionData)) {
                    value.value = actionData[key];
                }
            }

            // Creates the ticket in the RT API
            await this._createTicket(context, state);

            // Update the GUI to reflect the state of the ticket creation
            state.gui.buttons.create.enabled = false;
            // state.gui.buttons.cancel.title = "Borrar Hilo";
            // state.gui.buttons.cancel.tooltip = "Borra el hilo de conversacion";
            state.gui.buttons.cancel.title = "Ocultar ticket";
            state.gui.buttons.cancel.tooltip = "Oculta este ticket";

            const customFieldsJson = state.page1.body[4].items;
            for (const customFieldJson of customFieldsJson) {
                const keyJson: string = customFieldJson.items[1].items?.[0].id || customFieldJson.items[1].id; //customFieldJson.items[0].id;
                const cfState: any = state.ticket.customFields[keyJson];

                if (cfState.type === "Select") {
                    customFieldJson.items[1].items[0].type = "TextBlock";
                    customFieldJson.items[1].items[0].choices = [];
                    customFieldJson.items[1].selectAction.isEnabled = false;
                } else {
                    customFieldJson.items[1].text = cfState.value;
                    customFieldJson.items[1].type = "TextBlock";
                }
            }

            // Prepare the card data for the adaptive card
            const cardData: AdaptiveCardTicketCardPageData = {
                requestId: state.requestId,
                gui: state.gui,
            };

            // Expands the adaptive card template with the data provided
            const cardJson = new ACData.Template(page1).expand({
                $root: cardData,
            });

            // Creates a message attachment activity with the adaptive card using the expanded card template
            // and updated the existing adaptive card activity, id'ed by ' handlerContext.context.activity.replyToId'
            // with the new adaptive card JSON.
            const message = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
            message.id = context.activity.replyToId;

            // Sends the update to the existing adaptive card to the user
            await context.updateActivity(message);
        }
    }

    private async _createTicket(context: TurnContext, state: Record<string, any>): Promise<void> {
        const trigger: HandlerTriggerData = (context as any).trigger();

        // Get the initial message in the thread (The message that started the thread and contains a subject header)
        let thread: ChatMessage = trigger.thread!;

        let replies: ChatMessage[] = [];
        if (trigger.team?.aadGroupId && trigger.channel?.id && trigger.thread?.id) {
            // Get all the replies in the thread
            replies = await this._graph.teams
                .id(trigger.team!.aadGroupId!)
                .channels.id(trigger.channel!.id!)
                .messages.id(trigger.thread!.id!)
                .replies.get();
        }

        // Add the initial message to the replies and ticket description from the card to the beginning of the replies
        // to be added as comments to the ticket
        replies = [
            // {
            //   body: {
            //     content: handlerState.ticket.ticketDescriptionInput.value,
            //     contentType: "text/plain",
            //   },
            //   from: thread.from,
            // } as TeamChannelMessage,
            thread,
            ...replies,
        ];

        console.debug(`threadMessages.length: ${replies?.length}`);

        // Get the chosen queue from 'ticketCategoryChoiceSet' and get the queue from the API
        // const queue: Queue = await this._rt.queues.id(state.ticket.ticketCategoryChoiceSet.value).get();

        console.debug(`state.ticket:`, state.ticket);

        if (!trigger.threadFrom?.email) {
            throw new Error("Could not determine the email address of the user that started the thread");
        }
        if (!trigger.replyFrom?.email) {
            throw new Error("Could not determine the email address of the user that sent the reply");
        }

        const customFieldsBody: { [key: string]: string } = {};
        if (state.ticket.customFields && Object.keys(state.ticket.customFields).length > 0) {
            // If custom fields are provided, convert them to the expected format
            for (const [_, value] of Object.entries<any>(state.ticket.customFields)) {
                customFieldsBody[value.text] = value.value;
            }
        }

        // Create the ticket in the RT API
        // const ticket: Ticket = await this._rt.tickets.postTickets({
        //     requestBody: {
        //         Queue: queue.id,
        //         Subject: thread.subject ?? "No Subject",
        //         Status: state.ticket.ticketStateChoiceSet.value,
        //         Content: state.ticket.ticketDescriptionInput.value,
        //         TimeWorked: state.ticket.ticketTimeTakenInput.value,
        //         Requestor: trigger.threadFrom.email,
        //         Owner: trigger.replyFrom.email,
        //     } as any,
        // });
        // queue,
        // thread.subject ?? "No Subject",
        // state.ticket.ticketStateChoiceSet.value,
        // state.ticket.ticketTimeTakenInput.value,
        // state.ticket.ticketDescriptionInput.value,
        // trigger.threadFrom.email,
        // trigger.replyFrom.email,
        // state.ticket.customFields

        // const ticket: Ticket = await this._apiClient.ticket({
        //   id: "416134",
        //   _url: "https://test-epg-vmticket-01.hi.inet/REST/2.0/ticket/416134",
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

        // console.debug(`ticket:`, ticket);

        // Create a log entry for the ticket creation wuth the actionData and the thread messages that were
        // used to create the ticket
        // await this._logs.createLog(
        //   JSON.stringify({
        //     ...handlerState.ticket,
        //     threadMessages: replies,
        //   })
        // );

        // if (!ticket.id) {
        //     throw new Error("Ticket ID is undefined after ticket creation");
        // }

        for (const message of replies) {
            console.debug(`message.id: '${message.id}'`);

            if (
                !message ||
                typeof message !== "object" ||
                !("body" in message) ||
                typeof message.body !== "object" ||
                !message.body ||
                !("content" in message.body) ||
                typeof message.body.content !== "string" ||
                !message.body.content ||
                !message.body.content.trim() || // Should make sure trim exists right ts? no more trim() is not a function in 'message.body.content' error
                !message.from?.user
            ) {
                console.debug(`Message is not valid, skipping...`);
                continue;
            }

            if (
                "mentions" in message &&
                message.mentions &&
                Array.isArray(message.mentions) &&
                message.mentions.length === 1
            ) {
                const mention = message.mentions[0];
                const appMention = mention.mentioned?.application;

                if (
                    // appMention?.applicationIdentityType === ApplicationIdentityType.Bot &&
                    appMention?.id === this._config.botId
                ) {
                    // const messageText = message.body.content.toLowerCase();
                    // TODO: Message should only be skipped if it both mentions the bot and it evokes a known command or this command only?

                    console.debug(`Message mentions bot, skipping...`);
                    continue;
                }
            }

            // await this._rt.tickets.postTicketCorrespond({
            //     id: ticket.id,
            //     requestBody: {
            //         Subject: `Respuesta de ${message.from?.user?.displayName}`,
            //         ContentType: message.body.contentType || "text/plain",
            //         Content: message.body.content,
            //         Attachments: [],
            //         TimeTaken: "0",
            //     } as any,
            // });
            // addTicketComment(this._graph, ticket, message);
        }

        // Send a message to the user that the ticket was created and provide a link to the ticket
        // await context.sendActivity(
        //     `Se hay creado el ticket con el número: ${ticket.id}. Lo puedes acceder en [este enlace](${this._config.apiEndpoint}/Ticket/Display.html?id=${ticket.id}).`
        // );
    }
}

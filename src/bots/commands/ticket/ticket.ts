import { CardFactory, MessageFactory, TurnContext } from "botbuilder";

import * as ACData from "adaptivecards-templating";

import { HandlerTriggerData } from "../manager";
import { OAuthCommandHandler } from "../handler";
import { HandlerMessage } from "../message";
import { AdaptiveCardTicketCardPageData } from "../../adaptiveCards/actions/actions";
import { RTClient } from "../../../utils/client/rt/client";
import { Queue } from "../../../utils/client/rt/schemas/queues";

import page0 from "../../adaptiveCards/templates/ticket/page0.json";

interface QueueChoice {
    title: string;
    value: string;
}

export class TicketCommandHandler extends OAuthCommandHandler {
    private readonly Empty: string = " ";

    public pattern: RegExp = /\/?ticket/gi;

    constructor(private readonly _rt: RTClient) {
        super();
    }

    /**
     * @inheritdoc
     */
    public async run(context: TurnContext, message: HandlerMessage, _token?: string): Promise<void> {
        return await (context as any).switchToPersonalConversation(async (chatContext: TurnContext): Promise<void> => {
            return await this._run(chatContext, message, _token);
        });
    }

    private async _run(context: TurnContext, _message: HandlerMessage, _token?: string): Promise<void> {
        const data: HandlerTriggerData = (context as any).trigger();
        const state: Record<string, any> = (context as any).request().data || {};
        const requestId: string = (context as any).request().requestId;

        // Fetches the queue choices for the ticket
        const queueChoices: { title: string; value: string }[] = await this._fetchQueueChoices();

        // Fetches the status choices for the ticket
        const statusChoices: { title: string; value: string }[] = await this._fetchStatusChoices();

        // Creates a new ticket state in the handler state
        state.ticket = {
            startedAt: new Date(),
            ticketStateChoiceSet: {
                value: "",
                choices: statusChoices,
            },
            ticketCategoryChoiceSet: {
                value: "",
                choices: queueChoices,
                // required: false
            },
            ticketTimeTakenInput: {
                value: "0",
            },
            ticketDescriptionInput: {
                value: "",
            },
        };

        // Creates the 'startedAt' containing the time this card was sent in the 'es-ES' locale format
        const startedAt: string = state.ticket.startedAt.toLocaleString("es-ES", {
            timeZone: "UTC",
            month: "long",
            day: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
        });

        // Initializes the 'state.gui' object with the initial values for the adaptive card GUI state
        state.gui = {
            page: 0,
            header: {
                startedAt: startedAt,
                from: {
                    name: data?.replyFrom?.name ?? this.Empty,
                    email: data?.replyFrom?.email ?? this.Empty,
                },
            },
            context: {
                team: data?.team?.name ?? this.Empty,
                channel: data?.channel?.displayName ?? this.Empty,
                conversation: data?.thread?.subject ?? this.Empty,
            },
            buttons: {
                visible: true,
                create: {
                    title: "Siguiente",
                    tooltip: "Siguir con la creación de lo ticket",
                    enabled: true,
                },
                cancel: {
                    title: "Cancelar",
                    tooltip: "Cancela la creación de lo ticket",
                    enabled: true,
                },
            },
        };

        const cardData: AdaptiveCardTicketCardPageData = {
            requestId: requestId,
            ticket: state.ticket,
            gui: state.gui,
        };

        // Expands the adaptive card template with the data provided
        const cardJson = new ACData.Template(page0).expand({
            $root: cardData,
        });

        // Sends the adaptive card
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(cardJson)));
    }

    private async _fetchStatusChoices(): Promise<{ title: string; value: string }[]> {
        // Status choices array containing the title and value of each status to be displayed in the adaptive card
        return [
            { title: "Abierto", value: "open" },
            { title: "Nuevo", value: "new" },
            // { title: "Estancado", value: "stalled" },  // Not allowed in main queue
            { title: "Resuelto", value: "resolved" },
            // { title: "Rechazado", value: "rejected" }, // Not allowed in main queue
            // { title: "Borrado", value: "deleted" },    // Not allowed in main queue
        ];
    }

    private async _fetchQueueChoices(): Promise<QueueChoice[]> {
        // Queue choices array containing the title and value of each queue to be displayed in the adaptive card
        const queueChoices: QueueChoice[] = [];
        const queuesToChoices = (queues: Queue[]): QueueChoice[] => {
            return queues
                .map((queue: Queue) => {
                    if (queue?.id && queue?.Name) {
                        return { title: queue.Name, value: queue.id };
                    }
                    return null;
                })
                .filter((v) => v !== null);
        };

        await this._rt.queues.request.get().then(async (response) => {
            queueChoices.push(...queuesToChoices(response.items));
            while (response?.next_page) {
                response = await response.next().request.get();
                queueChoices.push(...queuesToChoices(response.items));
            }
        });

        // for (const queueRef of queues.items ?? []) {
        //     const queue: Queue | null = queueRef.id
        //         ? await this._rt.queues
        //               .id(queueRef.id)
        //               .request.get()
        //               .catch((error: any) => {
        //                   console.error(error);
        //                   return null;
        //               })
        //         : null;
        //     if (queue?.id && queue?.Name) {
        //         queueChoices.push({ title: queue.Name, value: queue.id });
        //     }
        // }

        // while (queues?.next_page) {
        // if (queues.page === undefined || queues.page === null || isNaN(queues.page)) {
        //     break;
        // }
        // const page: number = queues.page + 1;

        // queues = (await queues
        //     .next()
        //     .request.get()
        //     .catch((error: any) => {
        //         console.error(error);
        //         return null;
        //     }))!;

        // if (!queues) {
        //     break;
        // }

        //     for (const queueRef of queues?.items ?? []) {
        //         const queue: Queue | null = queueRef.id
        //             ? await this._rt.queues
        //                   .id(queueRef.id)
        //                   .request.get()
        //                   .catch((error: any) => {
        //                       console.error(error);
        //                       return null;
        //                   })
        //             : null;
        //         if (queue?.id && queue?.Name) {
        //             queueChoices.push({ title: queue.Name, value: queue.id });
        //         }
        //     }
        // }

        console.debug(`queueChoices:`, queueChoices);
        return queueChoices?.length > 0 ? queueChoices : [];
    }
}

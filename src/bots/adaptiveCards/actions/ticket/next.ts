import { CardFactory, MessageFactory, TurnContext } from "botbuilder";

import * as ACData from "adaptivecards-templating";

import { ActionHandler } from "../../../commands/handler";
import { HandlerMessage } from "../../../commands/message";
import {
    AdaptiveCardActionActivityValue,
    AdaptiveCardActionPositiveTicketPageData,
    AdaptiveCardTicketCardPageData,
} from "../actions";
import { isKeyOf } from "../../../../utils/misc";
import { RTClient } from "../../../../utils/client/rt/client";
import { CustomField } from "../../../../utils/client/rt/schemas/customFields";

import page1 from "../../templates/ticket/page1.json";
import customFieldSelect from "../../templates/ticket/customFieldSelect.json";
import customFieldInput from "../../templates/ticket/customFieldInput.json";

export class TicketAdaptiveCardNextActionHandler extends ActionHandler {
    public pattern: string = "nextTicket";

    constructor(private readonly _rt: RTClient) {
        super();
    }

    public async run(context: TurnContext, _message: HandlerMessage): Promise<any> {
        const activityValue: AdaptiveCardActionActivityValue = context.activity.value;
        const actionData: AdaptiveCardActionPositiveTicketPageData = activityValue?.action?.data;
        const data: Record<string, any> = (context as any).request().data;

        console.debug(`actionData:`, actionData);
        console.debug(`data:`, data);

        if (data.gui.page === 0) {
            // If the state is on page 0, we need to update the 'state.ticket' object with the data from the action
            for (const [key, field] of Object.entries<any>(data.ticket)) {
                if (isKeyOf(key, actionData)) {
                    field.value = actionData[key];
                }
            }

            const customFields: any[] = await this._buildCustomFields(data, actionData);
            page1.body[4].items = customFields;
            data.page1 = page1;

            customFields[0].items[1].items[0].choices =
                data.ticket.customFields[Number(customFields[0].items[1].items[0].id)].choices;
            customFields[0].items[1].selectAction.isEnabled = true;

            // Update the state GUI properties to reflect the state of the ticket creation
            data.gui.page = 1;
            data.gui.buttons.create.enabled = false;
            data.gui.buttons.create.title = "Crear Ticket";
            data.gui.buttons.create.tooltip = "Crea un nuevo ticket";

            // Prepare the card data for the adaptive card
            const cardData: AdaptiveCardTicketCardPageData = {
                requestId: actionData.requestId,
                gui: data.gui,
            };

            console.debug(`cardData:`, cardData);

            // Expands the adaptive card template with the data provided
            const cardJson = new ACData.Template(page1).expand({
                $root: cardData,
            });

            // console.debug(`cardJson:`, cardJson);

            // Update the existing adaptive card activity, id'ed by 'handlerContext.context.activity.replyToId'
            // with the new adaptive card JSON
            const message = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
            // message.id = handlerContext.context.activity.replyToId;
            // await handlerContext.context.updateActivity(message);

            const replyToId = context?.activity?.replyToId;
            if (replyToId) {
                await context.deleteActivity(replyToId);
            }
            await context.sendActivity(message);

            console.debug(`Moving to page 1`);
            return;
        }
        console.warn(`Next page ticket action invoke on unsupported page: ${data.gui.page}`);
        // TODO: Maybe throw
    }

    private async _buildCustomFields(
        state: Record<string, any>,
        actionData: AdaptiveCardActionPositiveTicketPageData,
    ): Promise<any[]> {
        if (!actionData.ticketCategoryChoiceSet) {
            return [];
        }
        const customFields: CustomField[] = await this._rt.queues
            .id(actionData.ticketCategoryChoiceSet)
            .ticketCustomFields.request.get();

        customFields.sort((a, b) => {
            if (a.id < b.id) {
                return -1;
            }
            if (a.id > b.id) {
                return 1;
            }
            return 0;
        });

        for (const customField of customFields) {
            if (customField.BasedOn) {
                for (const basedOnField of customFields) {
                    basedOnField.Dependents = basedOnField.Dependents || [];

                    if (basedOnField.id === customField.BasedOn.id) {
                        basedOnField.Dependents.push(customField);
                    }
                }
            }
        }

        const freeCustomFields: {
            [key: string]: {
                dependents: string[];
            };
        } = {};
        for (const field of customFields) {
            if (!field.BasedOn) {
                freeCustomFields[field.id] = {
                    dependents: field.Dependents?.map((cf: CustomField) => cf.id) ?? [],
                };
            }
        }

        const customFieldMap: {
            [key: string]: {
                id: string;
                text: string;
                placeholder: string;
                value: string;
                visible: boolean;
                basedOn: string | null;
                type: string;
                required: boolean;
                choices?: { value: string; title: string }[];
            };
        } = {};
        for (const customField of customFields) {
            customFieldMap[customField.id] = {
                id: customField.id,
                text: customField.Name,
                placeholder: customField.EntryHint ?? customField.Description ?? "",
                value: "",
                visible: customField.id in freeCustomFields,
                basedOn: customField.BasedOn?.id ?? null,
                type: customField.Type,
                required: customField.Type === "Select",
                choices:
                    customField.id in freeCustomFields
                        ? customField.Values?.map((choice: string, _index: number) => ({
                              value: choice,
                              title: choice,
                          }))
                        : [],
            };
        }
        state.ticket.customFields = customFieldMap;

        const customFieldJson: any[] = customFields.map((field: CustomField) => {
            if (!field.id) {
                return {};
            }

            if (field.Type === "Select") {
                return JSON.parse(
                    JSON.stringify(customFieldSelect)
                        .replace(/\${text}/g, customFieldMap[field.id].text)
                        .replace(/\${placeholder}/g, customFieldMap[field.id].placeholder)
                        .replace(/\${id}/g, customFieldMap[field.id].id)
                        .replace(/\${required}/g, String(field.id in freeCustomFields))
                        .replace(/\${requestId}/g, actionData.requestId)
                        .replace(/\${visible}/g, String(customFieldMap[field.id].visible))
                        .replace(/<id>/g, field.id),
                );
            } else if (field.Type === "Freeform") {
                return JSON.parse(
                    JSON.stringify(customFieldInput)
                        .replace(/\${text}/g, customFieldMap[field.id].text)
                        .replace(/\${placeholder}/g, customFieldMap[field.id].placeholder)
                        .replace(/\${id}/g, customFieldMap[field.id].id)
                        .replace(/\${requestId}/g, actionData.requestId)
                        .replace(/\${visible}/g, String(customFieldMap[field.id].visible))
                        .replace(/<id>/g, field.id),
                );
            }
        });

        return customFieldJson;
    }
}

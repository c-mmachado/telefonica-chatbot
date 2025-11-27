import { CardFactory, MessageFactory, TurnContext } from "botbuilder";

import * as ACTemplating from "adaptivecards-templating";

import {
    AdaptiveCardActionActivityValue,
    AdaptiveCardActionSelectChoiceData,
    AdaptiveCardTicketCardPageData,
} from "../actions";
import { ActionHandler } from "../../../commands/handler";
import { HandlerMessage } from "../../../commands/message";
// import { RTClient } from "../../../../utils/client/rt";

export class TicketAdaptiveCardSelectChoiceActionHandler implements ActionHandler {
    public pattern: string = "selectChoiceTicket";

    constructor() {} //private readonly _rt: RTClient

    public async run(context: TurnContext, _message: HandlerMessage): Promise<any> {
        const activityValue: AdaptiveCardActionActivityValue = context.activity.value;

        const actionData: AdaptiveCardActionSelectChoiceData = activityValue?.action?.data;

        // Calidate that we can retrieve the state
        const state: Record<string, any> = (context as any).request().data;
        if (!state) {
            throw new Error("Ticket adaptive card state is not initialized");
        }

        const page: number = state.gui.page;

        console.debug(`page:`, page);

        // For custom fields, we need to update the state with the selected choice
        const customFieldId: string = actionData.choice;
        await this._selectCustomFieldChoice(state, activityValue, customFieldId);

        // Check if all custom fields are filled in and if the create button should be enabled
        let enabled: boolean = true;
        const customFields = state.page1.body[4].items;
        for (const item of customFields) {
            const key = item.items[1].items?.[0].id || item.items[1].id; // item.items[0].id;
            const field = state.ticket.customFields[key];

            console.debug(`field.id:`, field.id, `field.value:`, field.value);

            if (field.required && !field.value && field.choices?.length > 0) {
                // If the field is required and has no value, we cannot enable the create button
                enabled = false;
                break;
            }
        }
        // Update the create button state based on the custom fields validation
        state.gui.buttons.create.enabled = enabled;
        console.debug(`enabled:`, enabled);

        // Update the GUI properties of the card to reflect the state of the ticket creation
        const cardData: AdaptiveCardTicketCardPageData = {
            requestId: state.requestId,
            gui: state.gui,
        };

        // Expands the adaptive card template with the data provided
        const cardJson = new ACTemplating.Template(state.page1).expand({
            $root: cardData,
        });

        console.debug(`cardJson:`, cardJson);

        // Update the card with the ticket information that was just submitted
        const message = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));

        // let textMessage = MessageFactory.text(" ");
        // textMessage.id = context.activity.replyToId;
        // await context.updateActivity(textMessage);
        // let invokeResponse = InvokeResponseFactory.adaptiveCard(
        //   CardFactory.adaptiveCard(cardJson)
        // );
        // await handlerContext.context.sendActivity({
        //   type: ActivityTypes.InvokeResponse,
        //   value: message,
        // } satisfies Partial<Activity>);

        message.id = context.activity.replyToId;
        await context.updateActivity(message);

        const replyToId = context?.activity?.replyToId;
        if (replyToId) {
            await context.deleteActivity(replyToId);
        }
        await context.sendActivity(message);

        // AdaptiveCards.updateCard(handlerContext.context, cardJson);

        return cardJson;
    }

    private async _selectCustomFieldChoice(
        state: Record<string, any>,
        activityValue: AdaptiveCardActionActivityValue,
        customFieldId: string
    ): Promise<void> {
        const customFieldValue: string = (activityValue as any)[customFieldId];
        const customFieldState = state.ticket.customFields[customFieldId];

        if (!customFieldValue) {
            // If customFieldValue is empty but customField.value is defined, it means the user wants to reset the field
            // If both are empty, we can just return
            if (!customFieldState?.value) {
                console.debug(`No change in custom field value, skipping reset...`);
                return;
            }

            const customFieldsJson = state.page1.body[4].items;
            for (const customFieldJson of customFieldsJson) {
                const keyJson: string = customFieldJson.items[1].items?.[0].id || customFieldJson.items[1].id; // customFieldJson.items[0].id;

                if (keyJson === customFieldId) {
                    console.debug(`Resetting field: '${keyJson}'`);

                    customFieldState.value = "";

                    if (customFieldJson.items[1].items) {
                        delete customFieldJson.items[1].items[0].text; // = ""

                        customFieldJson.items[1].items[0].type = "Input.ChoiceSet";
                        customFieldJson.items[1].items[0].value = "";
                        customFieldJson.items[1].items[0].placeholder = customFieldState.placeholder;
                        customFieldJson.items[1].items[0].choices = customFieldState.choices;
                        // customFieldJson.isVisible = true;
                        customFieldJson.items[1].items[0].isRequired = true;
                        customFieldJson.items[1].items[0].isMultiSelect = false;
                    }
                }
            }

            return await this._resetField(state, customFieldId, null);
        }

        if (!customFieldState) {
            throw new Error(`Custom field with id '${customFieldId}' not found in the ticket state.`);
        }

        if (customFieldState.value === customFieldValue) {
            console.debug(`No change in custom field value, skipping update...`);
            return;
        }

        const customFieldsJson = state.page1.body[4].items;
        for (const customFieldJson of customFieldsJson) {
            const keyJson: string = customFieldJson.items[1].items?.[0].id || customFieldJson.items[1].id; // customFieldJson.items[0].id;
            const currentFieldState: any = state.ticket.customFields[keyJson];

            // Update the field value in the state with the "auto" inputs returned by the adaptive card
            if (keyJson in activityValue) {
                currentFieldState.value = (activityValue as any)[keyJson] || "";
            }

            // Update the field value in the card
            // if (customFieldState.type === "Select") {
            //   customFieldJson.items[1].items[0].value = customFieldState.value;
            // } else {
            //   customFieldJson.items[1].value = customFieldState.value;
            // }

            if (keyJson === customFieldId) {
                console.debug(`Updating field: '${keyJson}'`);
                console.debug(`customFieldValue:`, customFieldJson);

                if (customFieldJson.items[1].items) {
                    // item.items[1].items[0].type = "TextBlock";
                    // item.items[1].items[0].text = customFieldValue;
                    // item.items[1].items[0].value = customFieldValue;
                    // item.items[1].items[0].isRequired = true;
                    customFieldJson.items[1].items[0].type = "TextBlock";
                    customFieldJson.items[1].items[0].text = currentFieldState.value;

                    delete customFieldJson.items[1].items[0].value; // = currentFieldState.value;
                    delete customFieldJson.items[1].items[0].isRequired; // = true
                    // customFieldJson.isVisible  = true
                    delete customFieldJson.items[1].items[0].choices;
                    delete customFieldJson.items[1].items[0].placeholder;
                    delete customFieldJson.items[1].items[0].isMultiSelect;
                }
            }
        }

        console.debug(`state.ticket:`, state.ticket);

        return await this._resetField(state, customFieldId, customFieldValue);
    }

    private async _resetField(
        state: Record<string, any>,
        customFieldId: string,
        customFieldValue: string | null
    ): Promise<void> {
        console.debug(`customFieldId: '${customFieldId}'`);

        // Once a field value changes all other fields that are 'basedOn' this field
        // should be reset to empty string and its choices should be recalculated
        for (const key of Object.keys(state.ticket.customFields)) {
            const customFieldState = state.ticket.customFields[key];

            console.debug(`field.id: '${customFieldState.id}', field.basedOn: '${customFieldState.basedOn}'`);

            if (customFieldState.basedOn === customFieldId) {
                console.debug(`Resetting field:`, customFieldState.id);

                let choices: { title: string; value: string }[] = [];
                if (customFieldValue) {
                    choices = []
                    // await this._rt
                    //     .customFieldValues(customFieldState.id, customFieldValue)
                    //     .then((response: CustomFieldValue[]) => {
                    //         return response.map((value: CustomFieldValue) => {
                    //             return { title: value.Name, value: value.Name };
                    //         });
                    //     });
                }

                console.debug(`choices:`, choices);

                customFieldState.value = "";
                customFieldState.choices = choices;

                const customFieldsJson = state.page1.body[4].items;
                for (const customFieldJson of customFieldsJson) {
                    const keyJson: string = customFieldJson.items[1].items?.[0].id || customFieldJson.items[1].id; // customFieldJson.items[0].id;

                    if (keyJson === String(customFieldState.id)) {
                        console.debug(`Updating field:`, keyJson);

                        delete customFieldJson.items[1].items[0].text; // = "";

                        customFieldJson.items[1].items[0].type = "Input.ChoiceSet";
                        customFieldJson.items[1].items[0].choices = choices;
                        customFieldJson.items[1].items[0].value = "";
                        customFieldJson.items[1].items[0].placeholder = customFieldState.placeholder;
                        customFieldJson.items[1].items[0].isRequired = choices.length > 0;
                        // customFieldJson.isVisible = choices.length > 0;
                        customFieldJson.items[1].selectAction.isEnabled = choices.length > 0;
                        break;
                    }
                }

                this._resetField(state, String(customFieldState.id), null);
            }
        }
    }
}

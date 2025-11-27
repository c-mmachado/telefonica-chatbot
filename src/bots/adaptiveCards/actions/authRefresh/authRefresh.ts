import { TurnContext } from "botbuilder";
import { DialogTurnResult } from "botbuilder-dialogs";

import { HandlerMessage } from "../../../commands/message";
import { ActionHandler } from "../../../commands/handler";
import { OAuthDialog } from "../../../dialogs/oauthDialog";
import { AdaptiveCardActionActivityValue, AdaptiveCardActionAuthRefreshDataOutput } from "../actions";

export class AuthRefreshActionHandler extends ActionHandler {
    /**
     * @inheritDoc
     */
    public pattern: string = "authRefresh";

    /**
     * @inheritDoc
     */
    public async run(context: TurnContext, _message: HandlerMessage): Promise<any> {
        if (context.activity.conversation.isGroup) {
            // This action should only ever be triggered in a personal context, do nothing if it isn't
            return;
        }

        const replyToId = context.activity?.replyToId;
        if (replyToId) {
            // Delete any previously sent message by the bot
            await context.deleteActivity(replyToId);
        }

        // Extract the data from the action
        const value: AdaptiveCardActionActivityValue = context.activity.value;
        const cardData: AdaptiveCardActionAuthRefreshDataOutput = value.action.data;

        console.debug(`cardData:`, cardData);

        // Run the dialog with the command and data from the action
        const dialogResult: DialogTurnResult = await (context as any).runDialog(context, OAuthDialog.name, {
            ...cardData,
        });

        console.debug(`dialogResult:`, dialogResult);
    }
}

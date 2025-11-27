import { TurnContext } from "botbuilder";

import { ActionHandler } from "../../../commands/handler";
import { HandlerMessage } from "../../../commands/message";

export class TicketAdaptiveCardCancelActionHandler implements ActionHandler {
    public pattern: string = "cancelTicket";

    constructor() {} // private readonly _graph: GraphClient

    public async run(context: TurnContext, _message: HandlerMessage): Promise<any> {
        // const activityValue: AdaptiveCardActionActivityValue =
        //   context.activity.value;

        // const actionData: any = activityValue?.action?.data;

        // const state: Record<string, any> = (context as any).request().data;

        // Deletes the adaptive card activity that triggered this action
        const replyToId = context.activity?.replyToId;
        if (replyToId) {
            await context.deleteActivity(replyToId);
        }
    }
}

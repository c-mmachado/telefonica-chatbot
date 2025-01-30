import { CommandMessage, TriggerPatterns } from "@microsoft/teamsfx";

import { ActionHandler, HandlerTurnContext } from "../../../commands/handler";

export class TicketAdaptiveCardCancelActionHandler implements ActionHandler {
  public pattern: TriggerPatterns = "cancelTicket";

  public async run(
    handlerContext: HandlerTurnContext,
    _: CommandMessage,
    __?: any
  ): Promise<any> {
    console.debug(
      `[${TicketAdaptiveCardCancelActionHandler.name}][DEBUG] [${this.run.name}]`
    );

    // Delete the ticket card
    await handlerContext.context.deleteActivity(
      handlerContext.context.activity.replyToId
    );
  }
}

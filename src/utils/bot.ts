import { CloudAdapter, TurnContext } from "botbuilder";
import {
  ClaimsIdentity,
  ConnectorClient,
  MicrosoftAppCredentials,
  TokenApiClient,
  UserTokenClient,
} from "botframework-connector";

export class BotUtils {
  public static getUserTokenClient(context: TurnContext): UserTokenClient {
    return context.turnState.get(
      (<CloudAdapter>context.adapter).UserTokenClientKey
    );
  }

  public static getMicrosoftAppCredentials(
    context: TurnContext
  ): MicrosoftAppCredentials {
    const userTokenClient = context.turnState.get(
      (<CloudAdapter>context.adapter).UserTokenClientKey
    );
    return (<TokenApiClient>userTokenClient["client"])
      .credentials as MicrosoftAppCredentials;
  }

  public static getBotIdentityKey(context: TurnContext): ClaimsIdentity {
    return context.turnState.get(
      (<CloudAdapter>context.adapter).BotIdentityKey
    );
  }

  public static getConnectorClientKey(context: TurnContext): ConnectorClient {
    return context.turnState.get(
      (<CloudAdapter>context.adapter).ConnectorClientKey
    );
  }
}

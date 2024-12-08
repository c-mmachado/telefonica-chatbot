import {
  AuthProviderCallback,
  Client,
} from "@microsoft/microsoft-graph-client";

export interface MicrosoftGraphEntity {
  "@odata.context": string;
}

export interface MicrosoftGraphCollection<T> extends MicrosoftGraphEntity {
  "@odata.count": number;
  "@odata.nextLink": string;
  value: T[];
}

export interface TeamChannelMessages
  extends MicrosoftGraphCollection<TeamChannelMessage> {}

export interface TeamChannelResponse extends MicrosoftGraphEntity {
  id: string;
  displayName: string;
  description: string;
  tenantId: string;
  isArchived: boolean;
}

export interface TeamChannelMessage extends MicrosoftGraphEntity {
  id: string;
  subject: string;
  attachments: TeamChannelMessageAttachment[];
  messageType: "message" | string;
  createdDateTime: Date;
  lastEditedDateTime: Date | null;
  deletedDateTime: Date | null;
  from: TeamChannelMessageFrom;
  webUrl: string;
}

export interface TeamChannelMessageFrom {
  id: string;
  tenantId: string;
  displayName: string;
  userIdentityType: string;
}

export interface TeamChannelMessageIdentity {
  teamId: string;
  channelId: string;
}

export interface TeamChannelMessageBody {
  content: string;
  contentType: string;
}

export interface TeamChannelMessageAttachment {
  id: string;
  content: any;
  contentType: string;
  contentUrl: string;
  name: string;
  teamsAppId: string;
  thumbnailUrl: string;
}

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
export class SimpleGraphClient {
  public static client(token: string): Client {
    return Client.init({
      authProvider: (done: AuthProviderCallback): void => {
        done(null, token);
      },
    });
  }

  public static async me(graphClient: Client): Promise<any> {
    return await graphClient
      .api("/me")
      .get()
      .catch((error) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.me.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async mePhoto(graphClient: Client): Promise<string> {
    try {
      const photo = await graphClient
        .api(`/me/photo/$value`)
        .get()
        .catch((error) => {
          console.error(
            `[${SimpleGraphClient.name}][ERROR] ${
              this.mePhoto.name
            } error:\n${JSON.stringify(error, null, 2)}`
          );
        });

      console.debug(
        `[${SimpleGraphClient.name}][DEBUG] [mePhoto] photo:\n${JSON.stringify(
          photo,
          null,
          2
        )}`
      );

      const arrayBuffer = await photo.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer, "binary");
      return "data:image/png;base64," + buffer.toString("base64");
    } catch (error: any) {
      console.error(
        `[${SimpleGraphClient.name}][ERROR] ${
          this.mePhoto.name
        } error:\n${JSON.stringify(error, null, 2)}`
      );
      return "";
    }
  }

  public static async teamChannel(
    graphClient: Client,
    teamAadGroupId: string,
    channelId: string
  ): Promise<TeamChannelResponse> {
    return await graphClient
      .api(`/teams/${teamAadGroupId}/channels/${channelId}`)
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamChannel.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async teamChannelMessage(
    graphClient: Client,
    teamAadGroupId: string,
    channelId: string,
    threadId: string
  ): Promise<TeamChannelMessage> {
    return await graphClient
      .api(
        `/teams/${teamAadGroupId}/channels/${channelId}/messages/${threadId}`
      )
      .version("beta")
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamChannelMessage.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async teamChannelMessages(
    graphClient: Client,
    teamAadGroupId: string,
    channelId: string,
    threadId: string
  ): Promise<TeamChannelMessages> {
    return await graphClient
      .api(
        `/teams/${teamAadGroupId}/channels/${channelId}/messages/${threadId}/replies`
      )
      .version("beta")
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamChannelMessages.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async teamChannelMessagesNext(
    graphClient: Client,
    odataNextLink: string
  ): Promise<TeamChannelMessages> {
    return await graphClient
      .api(odataNextLink)
      .version("beta")
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamChannelMessages.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }
}

// Init OnBehalfOfUserCredential instance with SSO token
// const oboCredential = new OnBehalfOfUserCredential(
//   token, // tokenResponse.ssoToken,
//   oboAuthConfig
// );

// // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
// const authProvider = new TokenCredentialAuthenticationProvider(
//   oboCredential,
//   {
//     scopes: [
//       "User.Read",
//       "Team.ReadBasic.All",
//       "Channel.ReadBasic.All",
//       "ChatMessage.Read",
//       "ProfilePhoto.Read.All",
//     ],
//   }
// );

// // Initialize Graph client instance with authProvider
// return Client.initWithMiddleware({
//   authProvider: authProvider,
// });

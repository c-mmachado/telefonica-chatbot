import {
  AuthProviderCallback,
  Client,
  ResponseType,
} from "@microsoft/microsoft-graph-client";

export enum ApplicationIdentityType {
  BOT = "bot",
}

export interface MicrosoftGraphEntity {
  "@odata.context": string;
}

export interface MicrosoftGraphCollection<T> extends MicrosoftGraphEntity {
  "@odata.count": number;
  "@odata.nextLink": string;
  value: T[];
}

export interface TeamsChannelMessages
  extends MicrosoftGraphCollection<TeamsChannelMessage> {}

export interface TeamsChannel extends MicrosoftGraphEntity {
  id: string;
  displayName: string;
  description: string;
  tenantId: string;
  isArchived: boolean;
}

export interface TeamsChannelMessage extends MicrosoftGraphEntity {
  id: string;
  subject: string;
  attachments: TeamsChannelMessageAttachment[];
  messageType: "message" | string;
  createdDateTime: Date;
  lastEditedDateTime: Date | null;
  deletedDateTime: Date | null;
  from: TeamsChannelMessageFrom;
  webUrl: string;
  body: TeamsChannelMessageBody;
  mentions: TeamsMessageMention[];
}

export interface TeamsMessageMention {
  id: number;
  mentionText: string;
  mentioned: {
    device: any | null; // Better typing
    user: {
      "@odata.type": string;
      id: string;
      displayName: string;
      userIdentityType: string;
      tenantId: string;
    } | null;
    conversation: any | null; // Better typing
    tag: any | null; // Better typing
    application: {
      "@odata.type": string;
      id: string;
      displayName: string;
      applicationIdentityType: string;
    } | null;
  };
}

export interface TeamsChannelMessageFrom {
  user: {
    id: string;
    tenantId: string;
    displayName: string;
    userIdentityType: string;
  };
}

export interface TeamsChannelMessageIdentity {
  teamId: string;
  channelId: string;
}

export interface TeamsChannelMessageBody {
  content: string;
  contentType: string;
}

export interface TeamsChannelMessageAttachment {
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

  public static async teamsChannel(
    graphClient: Client,
    teamAadGroupId: string,
    channelId: string
  ): Promise<TeamsChannel> {
    return await graphClient
      .api(`/teams/${teamAadGroupId}/channels/${channelId}`)
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamsChannel.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async teamsChannelMessage(
    graphClient: Client,
    teamAadGroupId: string,
    channelId: string,
    threadId: string
  ): Promise<TeamsChannelMessage> {
    return await graphClient
      .api(
        `/teams/${teamAadGroupId}/channels/${channelId}/messages/${threadId}`
      )
      .version("beta")
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamsChannelMessage.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async teamsChannelMessages(
    graphClient: Client,
    teamAadGroupId: string,
    channelId: string,
    threadId: string
  ): Promise<TeamsChannelMessages> {
    return await graphClient
      .api(
        `/teams/${teamAadGroupId}/channels/${channelId}/messages/${threadId}/replies`
      )
      .version("beta")
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamsChannelMessages.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async teamsChannelMessagesNext(
    graphClient: Client,
    odataNextLink: string
  ): Promise<TeamsChannelMessages> {
    return await graphClient
      .api(odataNextLink)
      .version("beta")
      .get()
      .catch((error: any) => {
        console.error(
          `[${SimpleGraphClient.name}][ERROR] ${
            this.teamsChannelMessagesNext.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
  }

  public static async downloadFile(
    graphClient: Client,
    url: string
  ): Promise<ArrayBuffer> {
    const buffer = graphClient
      .api(url)
      .responseType(ResponseType.ARRAYBUFFER)
      .get();

    const buffer2 = await fetch(url).then((res: Response) => res.arrayBuffer());

    console.debug(
      `[${SimpleGraphClient.name}][DEBUG] [${
        this.downloadFile.name
      }] buffer:\n${JSON.stringify(buffer, null, 2)}`
    );

    console.debug(
      `[${SimpleGraphClient.name}][DEBUG] [${
        this.downloadFile.name
      }] buffer2:\n${JSON.stringify(buffer2, null, 2)}`
    );
    return buffer;
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

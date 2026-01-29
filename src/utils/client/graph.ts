import { AuthenticationProvider, Client } from "@microsoft/microsoft-graph-client";
import { Channel, ChatMessage, Team, User } from "@microsoft/microsoft-graph-types";

import { AuthenticationResult, ConfidentialClientApplication } from "@azure/msal-node";
import { AuthenticationProviderOptions } from "@microsoft/microsoft-graph-client";

import { config } from "../../config/config";

export interface RequestBuilder<T> {
    get(): Promise<T>;
}

export interface GraphClient {
    teams: TeamsRequestBuilder;

    me(): Promise<User>;
}

export interface TeamsRequestBuilder extends RequestBuilder<Team[]> {
    id(teamId: string): TeamRequestBuilder;
}

export interface TeamRequestBuilder extends RequestBuilder<Team> {
    channels: ChannelsRequestBuilder;

    members: MembersRequestBuilder;

    id(teamId: string): TeamRequestBuilder;
}

export interface MembersRequestBuilder extends RequestBuilder<User[]> {
    id(userId: string): MemberRequestBuilder;
}

export interface MemberRequestBuilder extends RequestBuilder<User> {
    id(userId: string): MemberRequestBuilder;
}

export interface ChannelsRequestBuilder extends RequestBuilder<Channel[]> {
    id(channelId: string): ChannelRequestBuilder;
}

export interface ChannelRequestBuilder extends RequestBuilder<Channel> {
    messages: MessagesRequestBuilder;

    id(channelId: string): ChannelRequestBuilder;
}

export interface MessagesRequestBuilder extends RequestBuilder<ChatMessage[]> {
    id(messageId: string): MessageRequestBuilder;
}

export interface MessageRequestBuilder extends RequestBuilder<ChatMessage> {
    replies: RepliesRequestBuilder;

    id(messageId: string): MessageRequestBuilder;
}

export interface RepliesRequestBuilder extends RequestBuilder<ChatMessage[]> {
    id(messageId: string): ReplyRequestBuilder;
}

export interface ReplyRequestBuilder extends RequestBuilder<ChatMessage> {
    id(replyId: string): ReplyRequestBuilder;
}

// export enum ApplicationIdentityType {
//     Bot = "bot",
// }

export interface DefaultGraphClientOptions {
    authProvider: AuthenticationProvider;
}

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
export class DefaultGraphClient implements GraphClient {
    public static readonly DefaultScope = "https://graph.microsoft.com/.default";

    private readonly _client: Client;

    public teams: TeamsRequestBuilder;

    constructor(options: DefaultGraphClientOptions) {
        this._client = Client.initWithMiddleware({
            debugLogging: true,
            authProvider: options.authProvider,
        });
        this.teams = new DefaultTeamsRequestBuilder(this._client);
    }

    public async me(): Promise<User> {
        return this._client.api("/me").get();
    }
}

class DefaultTeamsRequestBuilder implements TeamsRequestBuilder {
    constructor(private readonly _client: Client) {}

    public id(teamId: string): TeamRequestBuilder {
        return new DefaultTeamRequestBuilder(this._client, teamId);
    }

    public async get(): Promise<Team[]> {
        return this._client.api("/teams").get();
    }
}

class DefaultTeamRequestBuilder implements TeamRequestBuilder {
    public channels: ChannelsRequestBuilder = new DefaultChannelsRequestBuilder(this._client, this._teamId);

    public members: MembersRequestBuilder = new DefaultMembersRequestBuilder(this._client, this._teamId);

    constructor(
        private readonly _client: Client,
        private _teamId: string,
    ) {}

    public id(teamId: string): TeamRequestBuilder {
        this._teamId = teamId;
        return this;
    }

    public async get(): Promise<Team> {
        return this._client.api(`/teams/${this._teamId}`).get();
    }
}

class DefaultMembersRequestBuilder implements MembersRequestBuilder {
    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
    ) {}

    public id(userId: string): MemberRequestBuilder {
        return new DefaultMemberRequestBuilder(this._client, this._teamId, userId);
    }

    public async get(): Promise<User[]> {
        return this._client.api(`/teams/${this._teamId}/members`).get();
    }
}

class DefaultMemberRequestBuilder implements MemberRequestBuilder {
    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
        private _userId: string,
    ) {}

    public id(userId: string): MemberRequestBuilder {
        this._userId = userId;
        return this;
    }

    public async get(): Promise<User> {
        return this._client.api(`/teams/${this._teamId}/members/${this._userId}`).get();
    }
}

class DefaultChannelsRequestBuilder implements ChannelsRequestBuilder {
    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
    ) {}

    public id(channelId: string): ChannelRequestBuilder {
        return new DefaultChannelRequestBuilder(this._client, this._teamId, channelId);
    }

    public async get(): Promise<Channel[]> {
        return this._client.api(`/teams/${this._teamId}/channels`).get();
    }
}

class DefaultChannelRequestBuilder implements ChannelRequestBuilder {
    public messages: MessagesRequestBuilder = new DefaultMessagesRequestBuilder(
        this._client,
        this._teamId,
        this._channelId,
    );

    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
        private _channelId: string,
    ) {}

    public id(channelId: string): ChannelRequestBuilder {
        this._channelId = channelId;
        return this;
    }

    public async get(): Promise<Channel> {
        return this._client.api(`/teams/${this._teamId}/channels/${this._channelId}`).get();
    }
}

class DefaultMessagesRequestBuilder implements MessagesRequestBuilder {
    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
        private readonly _channelId: string,
    ) {}

    public id(messageId: string): MessageRequestBuilder {
        return new DefaultMessageRequestBuilder(this._client, this._teamId, this._channelId, messageId);
    }

    public async get(): Promise<ChatMessage[]> {
        return this._client.api(`/teams/${this._teamId}/channels/${this._channelId}/messages`).get();
    }
}

class DefaultMessageRequestBuilder implements MessageRequestBuilder {
    public replies: RepliesRequestBuilder = new DefaultRepliesRequestBuilder(
        this._client,
        this._teamId,
        this._channelId,
        this._messageId,
    );

    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
        private readonly _channelId: string,
        private _messageId: string,
    ) {}

    public id(messageId: string): MessageRequestBuilder {
        this._messageId = messageId;
        return this;
    }

    public async get(): Promise<ChatMessage> {
        return this._client.api(`/teams/${this._teamId}/channels/${this._channelId}/messages/${this._messageId}`).get();
    }
}

class DefaultRepliesRequestBuilder implements RepliesRequestBuilder {
    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
        private readonly _channelId: string,
        private readonly _messageId: string,
    ) {}

    public id(replyId: string): ReplyRequestBuilder {
        return new DefaultReplyRequestBuilder(this._client, this._teamId, this._channelId, this._messageId, replyId);
    }

    public async get(): Promise<ChatMessage[]> {
        return this._client
            .api(`/teams/${this._teamId}/channels/${this._channelId}/messages/${this._messageId}/replies`)
            .get();
    }
}

class DefaultReplyRequestBuilder implements ReplyRequestBuilder {
    constructor(
        private readonly _client: Client,
        private readonly _teamId: string,
        private readonly _channelId: string,
        private readonly _messageId: string,
        private _replyId: string,
    ) {}

    public id(replyId: string): ReplyRequestBuilder {
        this._replyId = replyId;
        return this;
    }

    public async get(): Promise<ChatMessage> {
        return this._client
            .api(
                `/teams/${this._teamId}/channels/${this._channelId}/messages/${this._messageId}/replies/${this._replyId}`,
            )
            .get();
    }
}

export const msalClient = new ConfidentialClientApplication({
    auth: {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        authority: `${config.authorityHost}/${config.tenantId}`,
    },
});

export const graphClient: DefaultGraphClient = new DefaultGraphClient({
    authProvider: {
        async getAccessToken(
            _authenticationProviderOptions: AuthenticationProviderOptions | undefined,
        ): Promise<string> {
            msalClient.clearCache();
            const result: AuthenticationResult | null = await msalClient.acquireTokenByClientCredential({
                scopes: [DefaultGraphClient.DefaultScope], // "https://graph.microsoft.com/.default"
            });
            if (!result?.accessToken) {
                throw new Error("Could not acquire access token for Graph API client.");
            }
            // console.debug(result);
            return result.accessToken;
        },
    },
});

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

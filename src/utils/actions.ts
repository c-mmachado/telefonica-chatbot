import {
  ChannelAccount,
  ChannelInfo,
  ConversationAccount,
  TeamDetails,
} from "botbuilder";

export enum AdaptiveCardAction {
  Name = "adaptiveCard/action",

  AuthRefresh = "authRefresh",
  CreateTicket = "createTicket",
}

export type AdaptiveCardActionActivityValue = {
  action: {
    verb: string;
    data?: any & {
      command: string;
    };
  };
};

export type AdaptiveCardActionAuthRefreshDataInput = {
  command: string;
  team: TeamDetails;
  channel: ChannelInfo;
  conversation: ConversationAccount;
  from: ChannelAccount;
  userIds: string[];
};

export type AdaptiveCardActionAuthRefreshDataOutput = {
  command: string;
  team: TeamDetails;
  channel: ChannelInfo;
  conversation: ConversationAccount;
  from: ChannelAccount;
};

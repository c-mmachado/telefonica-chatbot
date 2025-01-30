import { Client } from "@microsoft/microsoft-graph-client";

import { BotConfiguration } from "../config/config";
import { SimpleGraphClient, TeamsChannelMessage } from "./graphClient";
import { AppInstallUtils } from "./appInstall";

type Required<T, U extends keyof T> = T & { [key in U]-?: T[key] };

export type HyperlinkEntity = Partial<{
  _url: string;
  ref: string;
  id: string;
  type: string;
  name: string;

  from: string;
  to: string;
  label: string;
  update: string;
}>;

export type TypedHyperlinkEntity = HyperlinkEntity &
  Required<HyperlinkEntity, "type" | "_url">;

export type RefHyperlinkEntity = HyperlinkEntity &
  Required<HyperlinkEntity, "ref" | "_url">;

export type CustomFieldHyperlink = TypedHyperlinkEntity &
  Required<HyperlinkEntity, "name"> & {
    values?: string[];
  };

export interface PagedCollection<T> {
  items: T[];
  next_page?: string;
  prev_page?: string;
  page: number;
  per_page: number;
  total: number;
  pages: number;
  count: number;
}

export interface TicketHistory extends PagedCollection<TypedHyperlinkEntity> {}

export interface Queues extends PagedCollection<TypedHyperlinkEntity> {}

export interface Queue {
  id: string;
  Name: string;
  _hyperlinks: RefHyperlinkEntity[];
}

export interface CreateTicket extends TypedHyperlinkEntity {}

export interface Ticket {
  id: string;

  Subject: string;
  Type: string;
  Status: string;

  Requestor: string[];

  InitialPriority: number;
  Priority: number;
  FinalPriority: number;

  TimeLeft: number;
  TimeWorked: number;
  TimeEstimated: number;

  Cc: string[];
  AdminCc: string[];

  Started: Date;
  Resolved: Date;
  Starts: Date;
  Due: Date;
  Created: Date;
  LastUpdated: Date;

  Queue: TypedHyperlinkEntity;
  Owner: TypedHyperlinkEntity;
  Creator: TypedHyperlinkEntity;
  LastUpdatedBy: TypedHyperlinkEntity;

  EffectiveId: TypedHyperlinkEntity;

  CustomFields: CustomFieldHyperlink[];

  _hyperlinks: RefHyperlinkEntity[];
}

export class APIClient {
  private _cookie: string;

  constructor(private readonly _config: BotConfiguration) {}

  public async login(): Promise<any> {
    console.debug(
      `[${APIClient.name}][DEBUG] ${this.login.name} endpoint: ${this._config.apiEndpoint}`
    );

    const cookie = await fetch(`${this._config.apiEndpoint}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Accept:
          "text/html,\
          application/xhtml+xml,\
          application/xml;0.9,\
          image/avif,\
          image/webp,\
          image/apng;q=0.9,\
          image/avif,\
          image/webp,\
          image/apng,\
          /;q=0.8,\
          application/signed-exchange;v=b3;q=0.7",
      },
      body: new URLSearchParams({
        user: this._config.apiUsername,
        pass: this._config.apiPassword,
        next: "7a73ae647301ce8bdff23044613b37a3",
      }),
    })
      .then((response: Response): string => {
        //   const helper = async function (
        //     array: ReadableStream<Uint8Array>
        //   ): Promise<string> {
        //     const utf8 = new TextDecoder("utf-8");
        //     let asString = "";
        //     for await (const chunk of array as any) {
        //       // Can use for-await starting ES 2018
        //       asString += utf8.decode(chunk);
        //     }
        //     return asString;
        //   };
        //   const asString = await helper(response.body);
        //   for (const header of response.headers.entries()) {
        //     console.debug(
        //       `[${APIClient.name}][DEBUG] ${this.login.name} header:\n${header}`
        //     );
        //   }

        const setCookies = response.headers.getSetCookie();
        if (setCookies) {
          return setCookies[0];
        }
        return null;
      })
      .catch((error: any) => {
        console.error(
          `[${APIClient.name}][ERROR] ${
            this.login.name
          } error:\n${JSON.stringify(error, null, 2)}`
        );
      });
    return cookie;
  }

  public async get<T>(url: string): Promise<T> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.get.name} endpoint: ${url}`
    );

    return fetch(url, {
      method: "GET",
      headers: {
        Cookie: this._cookie,
        Accept: "application/json",
      },
    }).then((response: Response): Promise<T> => {
      return response?.json();
    });
  }

  public async next<T>(
    page: PagedCollection<T>
  ): Promise<PagedCollection<T> | null> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.next.name} endpoint: ${page.next_page}`
    );

    if (!page.next_page) {
      return null;
    }

    return this.get<PagedCollection<T>>(page.next_page);
  }

  public async queues(): Promise<Queues> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.queues.name} endpoint: ${this._config.apiEndpoint}/REST/2.0/queues/all}`
    );

    return this.get<Queues>(`${this._config.apiEndpoint}/REST/2.0/queues/all`);
  }

  public async queue(queue: TypedHyperlinkEntity | string): Promise<Queue> {
    if (!this._cookie) {
      // If the cookie is not set, login to the API and set the cookie
      this._cookie = await this.login();
    }

    if (typeof queue === "string") {
      // If the queue is a string, convert it to a hyperlink entity
      queue = {
        id: queue,
        _url: `${this._config.apiEndpoint}/REST/2.0/queue/${queue}`,
        type: "queue",
      };
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.queue.name} endpoint: ${queue._url}`
    );

    if (queue.type !== "queue") {
      // If the queue is not of the expected type, throw an error
      throw new Error(
        `The supplied hyperlink of type '${queue.type}' is not of the expected type 'queue'.`
      );
    }

    // Fetch the queue
    return this.get<Queue>(`${queue._url}`);
  }

  public async createTicket(queue: Queue, subject: string): Promise<Ticket> {
    if (!this._cookie) {
      // If the cookie is not set, login to the API and set the cookie
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${
        this.createTicket.name
      } queue:\n${JSON.stringify(queue, null, 2)}`
    );

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.createTicket.name} endpoint: ${`${
        queue._hyperlinks.find((v: RefHyperlinkEntity) => v.ref === "create")
          ._url
      }`}`
    );

    const createTicket: CreateTicket = await fetch(
      `${
        queue._hyperlinks.find((v: RefHyperlinkEntity) => v.ref === "create")
          ._url
      }`,
      {
        method: "POST",
        headers: {
          Cookie: this._cookie,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          Subject: subject,
        }),
      }
    ).then((response: Response): Promise<CreateTicket> => {
      return response?.json();
    });

    return await this.ticket(createTicket);
  }

  public async ticket(ticket: TypedHyperlinkEntity): Promise<Ticket> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.ticket.name} endpoint: ${ticket._url}`
    );

    if (ticket.type !== "ticket") {
      throw new Error(
        `The supplied hyperlink of type '${ticket.type}' is not of the expected type 'ticket'.`
      );
    }

    return await this.get<Ticket>(ticket._url);
  }

  public async updateTicket(ticket: Ticket): Promise<any> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.updateTicket.name} endpoint: ${
        ticket._hyperlinks.find((v) => v.ref === "self")._url
      }`
    );

    return fetch(ticket._hyperlinks.find((v) => v.ref === "self")._url, {
      method: "PUT",
      headers: {
        Cookie: this._cookie,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        Status: ticket.Status,
      }),
    }).then((response: Response): Promise<any> => {
      return response?.json();
    });
  }

  public async addTicketComment(
    graphClient: Client,
    token: string,
    ticket: Partial<Ticket>,
    message: TeamsChannelMessage
  ): Promise<any> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${
        this.addTicketComment.name
      } message:\n${JSON.stringify(message, null, 2)}`
    );

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.addTicketComment.name} endpoint: ${
        ticket._hyperlinks.find((v) => v.ref === "comment")._url
      }`
    );

    // const botToken = await AppInstallUtils.getAccessToken(
    //   this._config.tenantId
    // );

    const attachments = [];
    if (message.attachments?.length > 0) {
      message.body.content += "<br><br>Attachments:<br>";

      for (const attachment of message.attachments) {
        message.body.content += `<a href="${attachment.contentUrl}">${attachment.name}</a><br>`;

        // const buffer = await fetch(attachment.contentUrl, {
        //   method: "GET",
        //   headers: {
        //     Authorization: `${botToken.token_type} ${botToken.access_token}`,
        //   },
        // }).then((res: Response) => {
        //   return res.arrayBuffer();
        // });

        // console.debug(
        //   `[${APIClient.name}][DEBUG] ${
        //     this.addTicketComment.name
        //   } buffer:\n${JSON.stringify(buffer, null, 2)}`
        // );

        // const decoder = new TextDecoder("utf-8");
        // const bufferStr = decoder.decode(buffer);
        // const encodedFile = Buffer.from(bufferStr, "binary").toString("base64");

        // console.debug(
        //   `[${APIClient.name}][DEBUG] ${
        //     this.addTicketComment.name
        //   } encodedFile:\n${JSON.stringify(encodedFile, null, 2)}`
        // );

        // attachments.push({
        //   FileName: attachment.id,
        //   FileType: "text/plain",
        //   FileContent: encodedFile,
        // });
      }
    }

    const createComment = await fetch(
      `${ticket._hyperlinks.find((v) => v.ref === "comment")._url}`,
      {
        method: "POST",
        headers: {
          Cookie: this._cookie,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          Subject: `Respuesta de ${message.from.user.displayName}`,
          Content: message.body.content,
          ContentType: "text/html",
          TimeTaken: "1",
          Attachments: attachments,
        }),
      }
    ).then((response: Response): Promise<any> => {
      return response?.json();
    });

    console.debug(
      `[${APIClient.name}][DEBUG] ${
        this.addTicketComment.name
      } createComment:\n${JSON.stringify(createComment, null, 2)}`
    );

    return createComment;
  }

  public async ticketHistory(ticket: Ticket): Promise<TicketHistory> {
    if (!this._cookie) {
      this._cookie = await this.login();
    }

    console.debug(
      `[${APIClient.name}][DEBUG] ${this.addTicketComment.name} endpoint: ${
        ticket._hyperlinks.find((v) => v.ref === "history")._url
      }`
    );

    const history: TicketHistory = await this.get<TicketHistory>(
      ticket._hyperlinks.find((v) => v.ref === "history")._url
    );

    console.debug(
      `[${APIClient.name}][DEBUG] ${
        this.ticketHistory.name
      } history:\n${JSON.stringify(history, null, 2)}`
    );

    return history;
  }
}

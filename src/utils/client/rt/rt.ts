import { HttpContentTypes, HttpHeaders, HttpMethods } from "../../../utils/http";
import { config } from "../../../config/config";
import {
    TypedHyperlinkEntity,
    RTConfig,
    QueuesConfig,
    QueueIdConfig,
    Queue,
    CustomField,
    CustomFieldValue,
    CustomFieldsConfig,
    CustomFieldConfig,
    CustomFieldValuesConfig,
    CustomFieldValueConfig,
    TicketsConfig,
    Ticket,
    CreateTicket,
    rtSchemaConfig,
    queuesSchemaConfig,
    queueSchemaConfig,
    customFieldsSchemaConfig,
    customFieldSchemaConfig,
    customFieldValuesSchemaConfig,
    customFieldValueSchemaConfig,
    ticketsSchemaConfig,
    ticketSchemaConfig,
    RTPagedCollection,
    RefHyperlinkEntity,
    QueueTicketCustomFieldsConfig,
    queueTicketCustomFieldsSchemaConfig,
    TicketIdConfig,
    QueueCustomFieldsConfig,
    queueCustomFieldsSchemaConfig,
    HyperlinkRef,
    TicketConfig,
    ticketIdSchemaConfig,
    QueueRef,
    createRTNavigatablePagedCollection,
} from "./schemas";
import {
    Client,
    createClient,
    BaseSchemaClientRequestBuilder,
    SchemaClientRequestBuilder,
    PagedSchemaClientRequestBuilder,
} from "../base";

export interface RTClient {
    rt: RootRequestBuilder;

    queues: QueuesRequestBuilder;

    customFields: CustomFieldsRequestBuilder;

    tickets: TicketsRequestBuilder;

    ticket: TicketRequestBuilder;
}

export interface RootRequestBuilder extends SchemaClientRequestBuilder<RTConfig> {}

export interface QueuesRequestBuilder extends PagedSchemaClientRequestBuilder<QueuesConfig> {
    id(queueId: string): QueueIdRequestBuilder;
}

export interface QueueIdRequestBuilder extends SchemaClientRequestBuilder<QueueIdConfig> {
    customFields: CustomFieldsRequestBuilder;

    ticketCustomFields: QueueTicketCustomFieldsRequestBuilder;
}

export interface QueueCustomFieldsRequestBuilder extends PagedSchemaClientRequestBuilder<QueueCustomFieldsConfig> {
    id(customFieldId: string): CustomFieldIdRequestBuilder;
}

export interface QueueTicketCustomFieldsRequestBuilder
    extends SchemaClientRequestBuilder<QueueTicketCustomFieldsConfig> {
    id(customFieldId: string): CustomFieldIdRequestBuilder;
}

export interface CustomFieldsRequestBuilder extends PagedSchemaClientRequestBuilder<CustomFieldsConfig> {
    id(customFieldId: string): CustomFieldIdRequestBuilder;
}

export interface CustomFieldIdRequestBuilder extends SchemaClientRequestBuilder<CustomFieldConfig> {
    customFieldValues: CustomFieldValuesRequestBuilder;
}

export interface CustomFieldValuesRequestBuilder extends PagedSchemaClientRequestBuilder<CustomFieldValuesConfig> {
    id(customFieldValueId: string): CustomFieldValueIdRequestBuilder;
}

export interface CustomFieldValueIdRequestBuilder extends SchemaClientRequestBuilder<CustomFieldValueConfig> {}

export interface TicketsRequestBuilder extends PagedSchemaClientRequestBuilder<TicketsConfig> {
    id(ticketId: string): TicketIdRequestBuilder;
}

export interface TicketRequestBuilder extends SchemaClientRequestBuilder<TicketConfig> {}

export interface TicketIdRequestBuilder extends SchemaClientRequestBuilder<TicketIdConfig> {}

export class DefaultRTClient implements RTClient {
    public readonly rt: RootRequestBuilder;

    public readonly queues: QueuesRequestBuilder;

    public readonly customFields: CustomFieldsRequestBuilder;

    public readonly tickets: TicketsRequestBuilder;

    public readonly ticket: TicketRequestBuilder;

    constructor(private readonly _client: Client) {
        this.rt = new DefaultRootRequestBuilder(this._client);
        this.queues = new DefaultQueuesRequestBuilder(this._client);
        this.customFields = new DefaultCustomFieldsRequestBuilder(this._client);
        this.tickets = new DefaultTicketsRequestBuilder(this._client);
        this.ticket = new DefaultTicketRequestBuilder(this._client);
    }
}

class DefaultRootRequestBuilder extends BaseSchemaClientRequestBuilder<RTConfig> implements RootRequestBuilder {
    constructor(client: Client) {
        super(client, rtSchemaConfig);
    }
}

class DefaultQueuesRequestBuilder extends BaseSchemaClientRequestBuilder<QueuesConfig> implements QueuesRequestBuilder {
    constructor(client: Client) {
        super(client, queuesSchemaConfig, {
            after: {
                get: async (response: unknown) => {
                    if (!response || typeof response !== "object") {
                        return Promise.reject(new Error("Invalid response format for queues"));
                    }
                    const page = response as RTPagedCollection<QueueRef>;
                    const refs = page.items ?? [];
                    const queues: Queue[] = await Promise.all(
                        refs
                            .filter((ref) => ref.ref === HyperlinkRef.Queue && ref.id)
                            .map((ref) => rt.queues.id(ref.id!).request.get())
                    );
                    const newPage: RTPagedCollection<Queue> = {
                        ...page,
                        items: queues,
                    };
                    return Promise.resolve(
                        createRTNavigatablePagedCollection<QueuesConfig, Queue>(
                            this.client,
                            queuesSchemaConfig,
                            newPage
                        )
                    );
                },
            },
        });
    }

    public id(queueId: string): QueueIdRequestBuilder {
        return new DefaultQueueIdRequestBuilder(this.client, queueId);
    }
}

class DefaultQueueIdRequestBuilder
    extends BaseSchemaClientRequestBuilder<QueueIdConfig>
    implements QueueIdRequestBuilder
{
    constructor(client: Client, private readonly _queueId: string) {
        super(client, queueSchemaConfig);
        this.variable("id", this._queueId);
    }

    public get customFields(): QueueCustomFieldsRequestBuilder {
        return new DefaultQueueCustomFieldsRequestBuilder(this.client);
    }

    public get ticketCustomFields(): QueueTicketCustomFieldsRequestBuilder {
        return new DefaultQueueTicketCustomFieldsRequestBuilder(this.client, this._queueId);
    }
}

class DefaultQueueCustomFieldsRequestBuilder
    extends BaseSchemaClientRequestBuilder<QueueCustomFieldsConfig>
    implements QueueCustomFieldsRequestBuilder
{
    constructor(client: Client) {
        super(client, queueCustomFieldsSchemaConfig);
    }

    public id(customFieldId: string): CustomFieldIdRequestBuilder {
        return new DefaultCustomFieldIdRequestBuilder(this.client, customFieldId);
    }
}

class DefaultQueueTicketCustomFieldsRequestBuilder
    extends BaseSchemaClientRequestBuilder<QueueTicketCustomFieldsConfig>
    implements QueueTicketCustomFieldsRequestBuilder
{
    constructor(client: Client, private readonly _queueId: string) {
        super(client, queueTicketCustomFieldsSchemaConfig, {
            after: {
                get: async (response: unknown): Promise<CustomField[]> => {
                    if (!response || typeof response !== "object") {
                        return Promise.reject(new Error("Invalid response format for queue ticket custom fields"));
                    }
                    const refs = (response as Queue).TicketCustomFields ?? [];
                    const fields = [];
                    for (const ref of refs) {
                        if (ref.ref === HyperlinkRef.CustomField && ref.id) {
                            fields.push(await rt.customFields.id(ref.id).request.get());
                        }
                    }
                    return Promise.resolve(fields);
                },
            },
        });
        this.variable("id", this._queueId);
    }

    public id(customFieldId: string): CustomFieldIdRequestBuilder {
        return new DefaultCustomFieldIdRequestBuilder(this.client, customFieldId);
    }
}

class DefaultCustomFieldsRequestBuilder
    extends BaseSchemaClientRequestBuilder<CustomFieldsConfig>
    implements CustomFieldsRequestBuilder
{
    constructor(client: Client) {
        super(client, customFieldsSchemaConfig);
    }

    public id(customFieldId: string): CustomFieldIdRequestBuilder {
        return new DefaultCustomFieldIdRequestBuilder(this.client, customFieldId);
    }
}

class DefaultCustomFieldIdRequestBuilder
    extends BaseSchemaClientRequestBuilder<CustomFieldConfig>
    implements CustomFieldIdRequestBuilder
{
    constructor(client: Client, private _customFieldId: string) {
        super(client, customFieldSchemaConfig);
        this.variable("id", this._customFieldId);
    }

    public get customFieldValues(): CustomFieldValuesRequestBuilder {
        return new DefaultCustomFieldValuesRequestBuilder(this.client, this._customFieldId);
    }
}

class DefaultCustomFieldValuesRequestBuilder
    extends BaseSchemaClientRequestBuilder<CustomFieldValuesConfig>
    implements CustomFieldValuesRequestBuilder
{
    constructor(client: Client, private readonly _customFieldId: string) {
        super(client, customFieldValuesSchemaConfig);
        this.variable("id", this._customFieldId);
    }

    public id(customFieldValueId: string): CustomFieldValueIdRequestBuilder {
        return new DefaultCustomFieldValueIdRequestBuilder(this.client, this._customFieldId, customFieldValueId);
    }
}

class DefaultCustomFieldValueIdRequestBuilder
    extends BaseSchemaClientRequestBuilder<CustomFieldValueConfig>
    implements CustomFieldValueIdRequestBuilder
{
    constructor(client: Client, private readonly _customFieldId: string, private readonly _customFieldValueId: string) {
        super(client, customFieldValueSchemaConfig);
        this.variable("id", this._customFieldId);
        this.variable("value_id", this._customFieldValueId);
    }
}

class DefaultTicketsRequestBuilder
    extends BaseSchemaClientRequestBuilder<TicketsConfig>
    implements TicketsRequestBuilder
{
    constructor(client: Client) {
        super(client, ticketsSchemaConfig);
    }

    public id(ticketId: string): TicketIdRequestBuilder {
        return new DefaultTicketIdRequestBuilder(this.client, ticketId);
    }
}

class DefaultTicketIdRequestBuilder
    extends BaseSchemaClientRequestBuilder<TicketIdConfig>
    implements TicketIdRequestBuilder
{
    // public readonly customFields: CustomFieldsRequestBuilder = new DefaultCustomFieldsRequestBuilder(
    //     this.client,
    //     `ticket/${this._ticket.id}`
    // );

    constructor(client: Client, ticketId: string) {
        super(client, ticketIdSchemaConfig);
        this.variable("id", ticketId);
    }
}

class DefaultTicketRequestBuilder extends BaseSchemaClientRequestBuilder<TicketConfig> implements TicketRequestBuilder {
    constructor(client: Client) {
        super(client, ticketSchemaConfig);
    }
}

export class APIClient {
    private _cookie: string | null = null;

    private readonly _path: string = `${this._config.apiEndpoint}${this._config.apiBasePath}`;

    constructor(private readonly _config: BotConfiguration) {}

    public async login(): Promise<string | null> {
        // Logs into the API and returns the cookie to be used in subsequent requests

        console.debug(`endpoint:`, this._config.apiEndpoint);

        // Fetch the cookie from the API using the supplied credentials
        const cookie: string | null = await fetch(`${this._config.apiEndpoint}`, {
            method: HttpMethods.Post,
            headers: {
                [HttpHeaders.ContentType]: HttpContentTypes.FormUrlEncoded,
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

                if (!response.ok) {
                    // If the response is not OK, throw an error to be caught by the catch block below
                    throw new Error(
                        `Failed to login to RT API. Status: ${response.status}, Status Text: ${response.statusText}`
                    );
                }

                // Gets the Set-Cookie header from the response
                const setCookies = response.headers.getSetCookie();
                if (setCookies) {
                    // If the Set-Cookie header is found, return the first cookie in the header (There can be multiple Set-Cookie headers)
                    return setCookies[0];
                }

                // If the Set-Cookie header is not found, throw an error to be caught by the catch block below
                throw new Error("No Set-Cookie header found in the response.");
            })
            .catch((error: any) => {
                console.error(error);
                return null;
            });

        // Returns the result of the login process to the caller
        return cookie;
    }

    public async get<T>(endpoint: string): Promise<T> {
        this._cookie = await this.login();
        if (!this._cookie) {
            throw new Error("Unable to login to the RT API.");
        }

        console.debug(`endpoint:`, endpoint);

        // Fetch the data from the API as a GET request
        return fetch(endpoint, {
            method: HttpMethods.Get,
            headers: {
                Cookie: this._cookie,
                Accept: HttpContentTypes.Json,
            },
        }).then((response: Response): Promise<T> => {
            // Return the JSON response from the API
            return response?.json();
        });
    }

    public async next<T>(page: RTPagedCollection<T>): Promise<RTPagedCollection<T> | null> {
        this._cookie = await this.login();
        if (!this._cookie) {
            throw new Error("Unable to login to the RT API.");
        }

        if (!page.next_page) {
            // If the next page is not found, return null
            return null;
        }

        console.debug(`endpoint:`, page.next_page);

        // Fetch the next page of the collection
        return this.get<RTPagedCollection<T>>(page.next_page);
    }

    public async queues(): Promise<QueuesPage> {
        const endpoint = `${this._path}/queues/all`;
        console.debug(`endpoint:`, endpoint);
        return this.get<QueuesPage>(endpoint);
    }

    public async queue(queue: TypedHyperlinkEntity | string): Promise<Queue> {
        if (typeof queue === "string") {
            // If the queue is a string, convert it to a hyperlink entity
            queue = {
                id: queue,
                _url: `${this._path}/queue/${queue}`,
                type: "queue",
            };
        }

        console.debug(`endpoint:`, queue._url);

        if (queue.type !== "queue") {
            // If the queue is not of the expected type, throw an error
            throw new Error(`The supplied hyperlink of type '${queue.type}' is not of the expected type 'queue'.`);
        }

        return this.get<Queue>(`${queue._url}`);
    }

    public async createTicket(
        queue: Queue,
        subject: string,
        status: string,
        timeWorked: string,
        description: string,
        requestorEmail: string,
        ownerEmail: string,
        customFields: {
            [key: string]: {
                text: string;
                value: string;
            };
        }
    ): Promise<Ticket> {
        console.debug(`queue:`, queue);

        const endpoint = queue?._hyperlinks.find((v: RefHyperlinkEntity) => v.ref === "create")?._url;

        if (!endpoint) {
            // If the `create` hyperlink is not found, throw an error
            throw new Error(`The supplied queue '${queue.id}' does not have a 'create' hyperlink.`);
        }

        console.debug(`endpoint:`, endpoint);

        this._cookie = await this.login();

        const customFieldsBody: { [key: string]: string } = {};
        if (customFields && Object.keys(customFields).length > 0) {
            // If custom fields are provided, convert them to the expected format
            for (const [_, value] of Object.entries(customFields)) {
                customFieldsBody[value.text] = value.value;
            }
        }

        const ownerUser: User = await this.user(ownerEmail);

        console.debug(`body:`, {
            Subject: subject,
            Status: status,
            Content: description,
            CustomFields: customFieldsBody,
            Requestor: requestorEmail,
            Owner: ownerUser?.Name,
            TimeWorked: timeWorked,
        });

        // Create the ticket using the supplied queue and subject
        const createTicket: CreateTicket = await fetch(endpoint, {
            method: HttpMethods.Post,
            headers: {
                Cookie: this._cookie!, // FIXME: Cookie is not guaranteed to be not null
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
            },
            body: JSON.stringify({
                Subject: subject,
                Status: status,
                Content: description,
                CustomFields: customFieldsBody,
                // Creator: creator,
                Requestor: requestorEmail,
                Owner: ownerUser?.Name,
            }),
        }).then((response: Response): Promise<CreateTicket> => {
            return response?.json();
        });

        console.debug(`createTicket response:`, createTicket);

        const ticket = await this.ticket(createTicket);

        if (
            timeWorked &&
            ((typeof timeWorked === "string" && !isNaN(parseInt(timeWorked, 10))) || typeof timeWorked == "number")
        ) {
            console.debug(`timeWorked:`, timeWorked);

            const selfEndpoint = ticket._hyperlinks.find((v) => v.ref === "self")?._url;
            if (!selfEndpoint) {
                throw new Error(`The supplied ticket '${ticket.id}' does not have a 'self' hyperlink`);
            }

            console.debug(`selfEndpoint:`, selfEndpoint);

            // If timeWorked is provided, update the ticket with the time worked
            const updateTicket = await this.updateTicket(selfEndpoint, {
                TimeWorked: timeWorked,
            }).catch((error: any) => {
                console.error(error);
            });

            console.debug(`updateTicket response:`, updateTicket);
        }

        return ticket;
    }

    public async ticket(ticket: TypedHyperlinkEntity): Promise<Ticket> {
        console.debug(`endpoint:`, ticket._url);

        if (ticket.type !== "ticket") {
            throw new Error(`The supplied hyperlink of type '${ticket.type}' is not of the expected type 'ticket'`);
        }

        return await this.get<Ticket>(ticket._url);
    }

    public async updateTicket(endpoint: string, body: any): Promise<any> {
        console.debug(`endpoint:`, endpoint);
        console.debug(`body:`, body);

        this._cookie = await this.login();

        return await fetch(endpoint, {
            method: HttpMethods.Put,
            headers: {
                Cookie: this._cookie!, // FIXME: Cookie is not guaranteed to be not null
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
            },
            body: JSON.stringify(body),
        }).then((response: Response): Promise<any> => {
            return response?.json();
        });
    }

    public async addTicketComment(
        _graphClient: GraphClient,
        ticket: Partial<Ticket>,
        message: ChatMessage
    ): Promise<string[]> {
        this._cookie = await this.login();

        console.debug(`message.id: '${message.id}'`);

        const endpoint = ticket._hyperlinks?.find((v) => v.ref === "comment")?._url;
        if (!endpoint) {
            // If the `comment` hyperlink is not found, throw an error
            throw new Error(`The supplied ticket '${ticket.id}' does not have a 'comment' hyperlink`);
        }

        console.debug(`endpoint:`, endpoint);

        // const botToken = await MicrosoftGraphUtils.getAccessToken(this._config);

        const attachments: any[] = [];
        if (message.body && (message.attachments?.length ?? 0) > 0) {
            message.body.content += "<br><br>Attachments:<br>";

            for (const attachment of message.attachments ?? []) {
                message.body.content += `<a href="${attachment.contentUrl}">${attachment.name}</a><br>`;

                // const buffer = await fetch(attachment.contentUrl, {
                //   method: "GET",
                //   headers: {
                //     Authorization: `${botToken.tokenType} ${botToken.token}`,
                //   },
                // }).then((res: Response): Promise<ArrayBuffer> => {
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

        const createComment: string[] = await fetch(endpoint, {
            method: HttpMethods.Post,
            headers: {
                Cookie: this._cookie!, // FIXME: Cookie is not guaranteed to be not null
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
            },
            body: JSON.stringify({
                Subject: `Respuesta de ${message.from?.user?.displayName}`,
                Content: message.body?.content,
                ContentType: HttpContentTypes.Html,
                TimeTaken: "0",
                Attachments: attachments,
            }),
        }).then((response: Response): Promise<string[]> => {
            return response?.json();
        });

        console.debug(`createComment response:`, createComment);

        return createComment;
    }

    public async ticketHistory(ticket: Partial<Ticket>): Promise<TicketTransaction> {
        const endpoint = ticket._hyperlinks?.find((v) => v.ref === "history")?._url;
        if (!endpoint) {
            throw new Error(`The supplied ticket '${ticket.id}' does not have a 'history' hyperlink.`);
        }

        console.debug(`history endpoint:`, endpoint);

        const history: TicketTransaction = await this.get<TicketTransaction>(endpoint);

        console.debug(`history:`, history);

        return history;
    }

    public async queueCustomFields(queueId: string): Promise<CustomField[]> {
        if (!queueId) {
            // If the queue ID is not provided, throw an error
            throw new Error("Argument 'queueId' is required to fetch custom fields");
        }

        const queue: Queue = await this.queue(queueId);

        const customFields: CustomField[] = [];
        if (!queue || !queue.TicketCustomFields) {
            // If the queue is not provided or it does not have 'TicketCustomFields', return an empty array
            return customFields;
        }

        for (const hyperlink of queue.TicketCustomFields) {
            if (hyperlink.ref === "customfield") {
                // If the hyperlink is of type 'customfield', fetch the custom fields
                const customField: CustomField = await this.get<CustomField>(hyperlink._url);

                if (!customField || !customField.id) {
                    // If the custom field is not found or does not have an ID, skip it
                    continue;
                }

                console.debug(`customField:`, customField);

                customFields.push(customField);
            }
        }

        return customFields;
    }

    public async customFieldValues(customFieldId: string, value: string): Promise<CustomFieldValue[]> {
        if (!customFieldId) {
            // If the custom field ID is not provided, throw an error
            throw new Error("Argument 'customFieldId' is required to fetch custom field values");
        }

        const endpoint = `${this._config.apiEndpoint}${this._config.apiBasePath}/customfield/${customFieldId}/values`;

        console.debug(`endpoint:`, endpoint);
        console.debug(`value:`, value);

        // Fetch the custom field choices
        return await this.get<RTPagedCollection<TypedHyperlinkEntity>>(endpoint).then(
            async (response: RTPagedCollection<TypedHyperlinkEntity>): Promise<CustomFieldValue[]> => {
                // Convert the response to an array of CustomFieldValue objects
                if (!response || response.items?.length === 0) {
                    // If the response is empty, return an empty array
                    return [];
                }

                // Map the response to CustomFieldValue objects
                const customFieldValues: CustomFieldValue[] = [];
                for (const ref of response.items) {
                    if (ref.type !== "customfieldvalue") {
                        continue;
                    }

                    const customFieldValue: CustomFieldValue = await this.get<CustomFieldValue>(ref._url);
                    if (!customFieldValue || !customFieldValue.id) {
                        // If the custom field value is not found or does not have an ID, skip it
                        continue;
                    }

                    if (customFieldValue.Category === value) {
                        customFieldValues.push(customFieldValue);
                    }
                }

                const total: number = response.total;
                let seen: number = response.count;
                let page: number = response.page;
                while (seen <= total) {
                    const nextResponse: RTPagedCollection<TypedHyperlinkEntity> = await this.get<
                        RTPagedCollection<TypedHyperlinkEntity>
                    >(
                        `${this._config.apiEndpoint}${
                            this._config.apiBasePath
                        }/customfield/${customFieldId}/values?page=${++page}`
                    );
                    if (!nextResponse || nextResponse.items?.length === 0) {
                        break;
                    }

                    seen += nextResponse.count;

                    console.debug(`seen: ${seen}`);

                    for (const ref of nextResponse.items) {
                        if (ref.type !== "customfieldvalue") {
                            continue;
                        }

                        const customFieldValue: CustomFieldValue = await this.get<CustomFieldValue>(ref._url);
                        if (!customFieldValue || !customFieldValue.id) {
                            // If the custom field value is not found or does not have an id, skip it
                            continue;
                        }

                        if (customFieldValue.Category === value) {
                            customFieldValues.push(customFieldValue);
                        }
                    }
                }

                return customFieldValues;
            }
        );
    }

    public async user(email: string): Promise<User> {
        this._cookie = await this.login();
        if (!this._cookie) {
            throw new Error("Unable to login to the RT API.");
        }

        console.debug(`email:`, email);

        // Fetch the user by email
        return await this.get<User>(`${this._config.apiEndpoint}${this._config.apiBasePath}/user/${email}`);
    }
}

const client = createClient(
    config.apiEndpoint,
    config.apiBasePath,
    async (): Promise<{ headerName: string; value: string }> => {
        return fetch(config.apiEndpoint, {
            method: HttpMethods.Post,
            headers: {
                [HttpHeaders.ContentType]: HttpContentTypes.FormUrlEncoded,
            },
            body: new URLSearchParams({
                user: config.apiUsername,
                pass: config.apiPassword,
                next: "7a73ae647301ce8bdff23044613b37a3",
            }),
        }).then(async (response: Response): Promise<{ headerName: string; value: string }> => {
            if (!response.ok) {
                throw new Error(`An error occurred while authenticating to the RT API. ${await response.text()}`);
            }

            const cookies = response.headers.getSetCookie();
            if (!cookies || cookies.length === 0) {
                throw new Error(
                    "No authentication cookie was returned by the RT API. The credentials may be invalid or there may be a server issue."
                );
            }

            return { headerName: HttpHeaders.Cookie, value: cookies[0] };
        });
    }
);
export const rt: RTClient = new DefaultRTClient(client);

rt.rt.request.get().then((response) => {});

rt.ticket.request
    .queryParam("Queue", "1")
    .post({
        Subject: "Test Ticket from API Client",
        Status: "new",
        Requestor: "user@example.com",
        Owner: "owner@example.com",
        Description: "This is a test ticket created using the RT API Client.",
        CustomFields: {
            Priority: "High",
        },
    })
    .then((response) => {});

rt.tickets.request.get().then((response) => {});

rt.tickets
    .id("1")
    .request.get()
    .then((response) => {});

rt.tickets
    .id("1")
    .request.put({
        Status: "open",
    })
    .then((response) => {});

rt.customFields.request.get().then((response) => {});

rt.customFields
    .id("1")
    .request.get()
    .then((response) => {});

rt.customFields
    .id("1")
    .customFieldValues.request.get()
    .then((response) => {});

rt.customFields
    .id("1")
    .customFieldValues.id("1")
    .request.get()
    .then((response) => {});

rt.queues.request.get().then((response) => {
    response
        .next()
        .request.get()
        .then((nextResponse) => {
            nextResponse?.items[0].id;
        });
});

rt.queues
    .id("1")
    .request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .customFields.request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .customFields.id("1")
    .request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .customFields.id("1")
    .customFieldValues.request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .customFields.id("1")
    .customFieldValues.id("1")
    .request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .ticketCustomFields.request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .ticketCustomFields.id("1")
    .request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .ticketCustomFields.id("1")
    .customFieldValues.request.get()
    .then((response) => {});

rt.queues
    .id("1")
    .ticketCustomFields.id("1")
    .customFieldValues.id("1")
    .request.get()
    .then((response) => {});

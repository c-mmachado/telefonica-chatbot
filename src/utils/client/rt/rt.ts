// import { ChatMessage } from "@microsoft/microsoft-graph-types";

// import { BotConfiguration } from "../../config/config";
// import { HttpContentTypes, HttpHeaders, HttpMethods } from "../http";

// import { GraphClient } from "./graph";

import z, { ZodSchema } from "zod";

import { HttpContentTypes, HttpHeaders, HttpMethods } from "../../../utils/http";
import { config } from "../../../config/config";
import { RequestBuilder } from "../common";
import {
    PagedCollection,
    TypedHyperlinkEntity,
    Ticket,
    CreateTicket,
    UpdateTicket,
    Queue,
    CustomField,
    QueueRef,
    CustomFieldValue,
    CustomFieldValueRef,
    TicketRef,
    CustomFieldRef,
    rtSchemaConfig,
    RTRoot,
} from "./model";

export interface Client {
    api(endpoint: string): ClientRequest;
}

export type SchemaConfig = {
    getResponse?: ZodSchema<any>;
    postRequest?: ZodSchema<any>;
    postResponse?: ZodSchema<any>;
    putRequest?: ZodSchema<any>;
    putResponse?: ZodSchema<any>;
    deleteResponse?: ZodSchema<any>;
};

type SchemaOrUndefined<T> = T extends ZodSchema ? z.infer<T> : undefined;

type InferFromConfig<C extends SchemaConfig> = {
    GetResponse: SchemaOrUndefined<C["getResponse"]>;
    PostRequest: SchemaOrUndefined<C["postRequest"]>;
    PostResponse: SchemaOrUndefined<C["postResponse"]>;
    PutRequest: SchemaOrUndefined<C["putRequest"]>;
    PutResponse: SchemaOrUndefined<C["putResponse"]>;
    DeleteResponse: SchemaOrUndefined<C["deleteResponse"]>;
};

// type ConditionalMethod<T, MethodName extends string> = T extends undefined
//     ? never
//     : { [K in MethodName]: () => Promise<T> };

export type Header = string | number | boolean | Array<string | number | boolean>;
export type Headers = Record<string, Header>;

export type QueryParam = string | number | boolean | Array<string | number | boolean>;
export type QueryParams = Record<string, QueryParam>;

export interface ClientRequest {
    get<GetResponse>(): Promise<GetResponse>;

    post<PostResponse>(content: unknown): Promise<PostResponse>;

    put<PutResponse>(content: unknown): Promise<PutResponse>;

    delete<DeleteResponse>(): Promise<DeleteResponse>;

    queryParam(name: string, value: QueryParam): ClientRequest;

    queryParams(params: QueryParams): ClientRequest;

    header(name: string, value: Header): ClientRequest;

    headers(headers: Headers): ClientRequest;

    // body(content: any): RTRequest<GetResponse, PostRequest, PostResponse, PutRequest, PutResponse, DeleteResponse>;
}

export type SchemaClientRequest<C extends SchemaConfig> = (InferFromConfig<C>["GetResponse"] extends undefined
    ? {}
    : {
          get: () => Promise<InferFromConfig<C>["GetResponse"]>;
      }) &
    (InferFromConfig<C>["PostResponse"] extends undefined
        ? {}
        : {
              post: (content: InferFromConfig<C>["PostRequest"]) => Promise<InferFromConfig<C>["PostResponse"]>;
          }) &
    (InferFromConfig<C>["PutResponse"] extends undefined
        ? {}
        : {
              put: (content: InferFromConfig<C>["PutRequest"]) => Promise<InferFromConfig<C>["PutResponse"]>;
          }) &
    (InferFromConfig<C>["DeleteResponse"] extends undefined
        ? {}
        : {
              delete: () => Promise<InferFromConfig<C>["DeleteResponse"]>;
          }) & {
        queryParam(name: string, value: QueryParam): SchemaClientRequest<C>;
        queryParams(params: QueryParams): SchemaClientRequest<C>;
        header(name: string, value: Header): SchemaClientRequest<C>;
        headers(headers: Headers): SchemaClientRequest<C>;
    };

// export type TypedRTRequest<C extends SchemaConfig> = {
// extends RTRequest<
//     InferFromConfig<C>["GetResponse"],
//     InferFromConfig<C>["PostRequest"],
//     InferFromConfig<C>["PostResponse"],
//     InferFromConfig<C>["PutRequest"],
//     InferFromConfig<C>["PutResponse"],
//     InferFromConfig<C>["DeleteResponse"]
// >
// get: InferFromConfig<C>["GetResponse"] extends undefined ? never : () => Promise<InferFromConfig<C>["GetResponse"]>;
// get: ConditionalMethod<InferFromConfig<C>["GetResponse"], "get">["get"];
// get(): Promise<InferFromConfig<C>["GetResponse"]>;

// post: InferFromConfig<C>["PostResponse"] extends undefined
//     ? never
//     : (body: InferFromConfig<C>["PostRequest"]) => Promise<InferFromConfig<C>["PostResponse"]>;
// post: ConditionalMethod<InferFromConfig<C>["PostResponse"], "post">["post"] extends infer U
//     ? U extends undefined
//         ? never
//         : (body: InferFromConfig<C>["PostRequest"]) => Promise<InferFromConfig<C>["PostResponse"]>
//     : never;
// post(content: InferFromConfig<C>["PostRequest"]): Promise<InferFromConfig<C>["PostResponse"]>;

// put: InferFromConfig<C>["PutResponse"] extends undefined
//     ? never
//     : (body: InferFromConfig<C>["PutRequest"]) => Promise<InferFromConfig<C>["PutResponse"]>;
// put: ConditionalMethod<InferFromConfig<C>["PutResponse"], "put">["put"] extends infer U
//     ? U extends undefined
//         ? never
//         : (id: string, body: InferFromConfig<C>["PutRequest"]) => Promise<InferFromConfig<C>["PutResponse"]>
//     : never;
// put(content: InferFromConfig<C>["PutRequest"]): Promise<InferFromConfig<C>["PutResponse"]>;

// delete: InferFromConfig<C>["DeleteResponse"] extends undefined
//     ? never
//     : () => Promise<InferFromConfig<C>["DeleteResponse"]>;
// delete: ConditionalMethod<InferFromConfig<C>["DeleteResponse"], "delete">["delete"];
// delete(): Promise<InferFromConfig<C>["DeleteResponse"]>;
// };

export type PagedCollectionRequest<C extends SchemaConfig> = {} & SchemaClientRequest<C>;

export interface PagedCollectionRequestBuilder<
    T extends DefaultPagedCollectionRequest<K>,
    K extends TypedHyperlinkEntity = TypedHyperlinkEntity
> extends RequestBuilder<T> {
    all: Promise<K[]>;
}

export interface RTClient {
    tickets: TicketsRequestBuilder;

    queues: QueuesRequestBuilder;

    rt: RootRequestBuilder;
}

export interface QueuesRequestBuilder extends PagedCollectionRequestBuilder<DefaultPagedCollectionRequest<QueueRef>> {
    id(queueId: string): QueueRequestBuilder;
}

export interface RootRequestBuilder extends RequestBuilder<RTRoot> {}

export interface QueueRequestBuilder extends RequestBuilder<Queue> {
    ticketCustomFields: CustomFieldsRequestBuilder;

    id(queueId: string): QueueRequestBuilder;
}

export interface CustomFieldsRequestBuilder
    extends PagedCollectionRequestBuilder<DefaultPagedCollectionRequest<CustomFieldRef>> {
    id(customFieldId: string): CustomFieldRequestBuilder;
}

export interface CustomFieldRequestBuilder extends RequestBuilder<CustomField> {
    customFieldValues: CustomFieldValuesRequestBuilder;

    id(customFieldId: string): CustomFieldRequestBuilder;
}

export interface CustomFieldValuesRequestBuilder
    extends PagedCollectionRequestBuilder<DefaultPagedCollectionRequest<CustomFieldValueRef>> {
    id(customFieldValueId: string): CustomFieldValueRequestBuilder;
}

export interface CustomFieldValueRequestBuilder extends RequestBuilder<CustomFieldValue> {
    id(customFieldValueId: string): CustomFieldValueRequestBuilder;
}

export interface TicketsRequestBuilder extends PagedCollectionRequestBuilder<DefaultPagedCollectionRequest<TicketRef>> {
    id(ticketId: string): TicketRequestBuilder;
}

export interface TicketRequestBuilder extends RequestBuilder<Ticket> {
    // history: TicketTransactionRequestBuilder;

    customFields: CustomFieldsRequestBuilder;

    id(ticketId: string): TicketRequestBuilder;

    subject(subject: string): TicketRequestBuilder;

    status(status: string): TicketRequestBuilder;

    description(description: string): TicketRequestBuilder;

    requestor(email: string): TicketRequestBuilder;

    owner(email: string): TicketRequestBuilder;

    customField(fieldName: string, value: string): TicketRequestBuilder;

    ticket(ticket: Partial<Ticket>): TicketRequestBuilder;

    create(): Promise<CreateTicket>;

    save(): Promise<UpdateTicket>;
}

// export interface TicketTransactionsRequestBuilder extends PagedCollectionRequestBuilder<TicketTransactionsRefPage> {
//     id(transactionId: string): TicketTransactionRequestBuilder;
// }

// export interface TicketTransactionRequestBuilder extends RequestBuilder<TicketTransaction> {
//     id(transactionId: string): TicketTransactionRequestBuilder;
// }

class DefaultClient implements Client {
    private readonly _endpoint: string;

    constructor(
        endpoint: string,
        basePath: string,
        private readonly _authProvider: () => Promise<{ headerName: string; value: string }>
    ) {
        this._endpoint = `${endpoint}${endpoint.endsWith("/") ? "" : "/"}${
            basePath.startsWith("/") ? basePath.slice(1) : basePath
        }${basePath.endsWith("/") ? "" : "/"}`;
    }

    public api(path: string): ClientRequest {
        return new DefaultClientRequest(
            `${this._endpoint}${path.startsWith("/") ? path.slice(1) : path}`,
            this._authProvider
        );
    }
}

class DefaultClientRequest implements ClientRequest {
    constructor(
        private readonly _endpoint: string,
        private readonly _authProvider: () => Promise<{ headerName: string; value: string }>,
        private readonly _queryParams: QueryParams = {},
        private readonly _headers: Headers = {}
    ) {}

    private _url(queryParams: QueryParams): URL {
        const url = new URL(this._endpoint);

        Object.entries(queryParams).forEach(([key, value]: [string, QueryParam]): void => {
            if (Array.isArray(value)) {
                value.forEach((val) => url.searchParams.append(key, String(val)));
            } else {
                url.searchParams.set(key, String(value));
            }
        });
        return url;
    }

    private _toHeaders(headers: Headers): HeadersInit {
        return Object.keys(headers).reduce((acc: Record<string, string>, key: string): Record<string, string> => {
            acc[key] = String(headers[key]);
            return acc;
        }, {});
    }

    private async _request(
        method: HttpMethods,
        options?: { body?: unknown; headers?: Headers; queryParams?: QueryParams }
    ): Promise<any> {
        const auth = await this._authProvider();
        return fetch(this._url(options?.queryParams ?? {}), {
            method: method,
            headers: {
                Accept: HttpContentTypes.Json,
                ...this._toHeaders(options?.headers ?? {}),
                [auth.headerName]: auth.value,
            },
            body: options?.body ? JSON.stringify(options?.body) : undefined,
        }).then((response: Response): Promise<any> => {
            if (!response.ok) {
                throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
            }
            return response.json();
        });
    }

    public async get<GetResponse>(): Promise<GetResponse> {
        return this._request(HttpMethods.Get, {
            headers: {
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public async post<PostResponse>(content?: unknown): Promise<PostResponse> {
        return this._request(HttpMethods.Post, {
            body: content,
            headers: {
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public async put<PutResponse>(content: unknown): Promise<PutResponse> {
        return this._request(HttpMethods.Put, {
            body: content,
            headers: {
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public async delete<DeleteResponse>(): Promise<DeleteResponse> {
        return this._request(HttpMethods.Delete, {
            headers: {
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public queryParam(name: string, value: QueryParam): ClientRequest {
        this._queryParams[name] = value;
        return this;
    }

    public queryParams(params: QueryParams): ClientRequest {
        Object.keys(params).forEach((key: string): void => {
            this._queryParams[key] = params[key];
        });
        return this;
    }

    public header(name: string, value: Header): ClientRequest {
        this._headers[name] = value;
        return this;
    }

    public headers(headers: Headers): ClientRequest {
        Object.keys(headers).forEach((key: string): void => {
            this._headers[key] = headers[key];
        });
        return this;
    }
}

class DefaultSchemaClientRequest<C extends SchemaConfig> {
    public static create<C extends SchemaConfig>(request: ClientRequest, config: C): SchemaClientRequest<C> {
        return new DefaultSchemaClientRequest<C>(request, config);
    }

    private constructor(private readonly _request: ClientRequest, private readonly _config: C) {}

    public async get(): Promise<InferFromConfig<C>["GetResponse"]> {
        if (!this._config.getResponse) {
            throw new Error("Resource does not support 'GET'");
        }
        const response = await this._request.get();
        return this._config.getResponse.parse(response);
    }

    public async post(content: InferFromConfig<C>["PostRequest"]): Promise<InferFromConfig<C>["PostResponse"]> {
        if (!this._config.postRequest || !this._config.postResponse) {
            throw new Error("Resource does not support 'POST'");
        }
        const validatedContent = this._config.postRequest?.parse(content);
        const response = await this._request.post(validatedContent);
        return this._config.postResponse.parse(response);
    }

    public async put(content: InferFromConfig<C>["PutRequest"]): Promise<InferFromConfig<C>["PutResponse"]> {
        if (!this._config.putRequest || !this._config.putResponse) {
            throw new Error("Resource does not support 'PUT'");
        }
        const validatedContent = this._config.putRequest?.parse(content);
        const response = await this._request.put(validatedContent);
        return this._config.putResponse.parse(response);
    }

    public async delete(): Promise<InferFromConfig<C>["DeleteResponse"]> {
        if (!this._config.deleteResponse) {
            throw new Error("Resource does not support 'DELETE'");
        }
        const response = await this._request.delete();
        return this._config.deleteResponse.parse(response);
    }

    public queryParam(name: string, value: QueryParam): SchemaClientRequest<C> {
        this._request.queryParam(name, value);
        return this;
    }

    public queryParams(params: QueryParams): SchemaClientRequest<C> {
        this._request.queryParams(params);
        return this;
    }

    public header(name: string, value: Header): SchemaClientRequest<C> {
        this._request.header(name, value);
        return this;
    }

    public headers(headers: Headers): SchemaClientRequest<C> {
        this._request.headers(headers);
        return this;
    }
}

export function createSchemaClientRequest<C extends SchemaConfig>(
    request: ClientRequest,
    config: C
): SchemaClientRequest<C> {
    const base = DefaultSchemaClientRequest.create<C>(request, config);
    const isEnabled: Record<"get" | "post" | "put" | "delete", boolean> = {
        get: !!config.getResponse,
        post: !!(config.postRequest && config.postResponse),
        put: !!(config.putRequest && config.putResponse),
        delete: !!config.deleteResponse,
    };

    // Wrap in Proxy to intercept access to disabled methods
    return new Proxy(base, {
        get(target: DefaultSchemaClientRequest<C>, propertyName: string | symbol, receiver: any): any {
            if (propertyName in base) {
                let keyName = propertyName as keyof typeof isEnabled;
                if (!isEnabled[keyName]) {
                    throw new Error(`Attempting to access disabled property/method '${String(propertyName)}'`);
                }

                const baseKeyName = keyName as typeof keyName & keyof typeof base;
                const value = base[baseKeyName];
                if (typeof value === "function") {
                    return value.bind(base);
                }
                return value;
            }

            const value = Reflect.get(target, propertyName, receiver);
            if (typeof value === "function") {
                return value.bind(target);
            }
            return value;
        },
    }) as SchemaClientRequest<C>;
}

// function createTypedRequest<C extends SchemaConfig>(request: RTRequest, config: C): new () => TypedRTRequest<C> {
//     return class ConfiguredRequest extends DefaultTypedRTRequest<C> {
//         constructor() {
//             super(request, config);
//         }
//     };
// }


export class DefaultPagedCollectionRequest<T extends TypedHyperlinkEntity> {
    constructor(private readonly _client: Client, private readonly _page: PagedCollection<T>) {}

    public get count(): number {
        return this._page.count;
    }

    public get total(): number {
        return this._page.total;
    }

    public get page(): number {
        return this._page.page;
    }

    public get per_page(): number {
        return this._page.per_page;
    }

    public get pages(): number {
        return this._page.pages;
    }

    public get items(): T[] {
        return this._page.items;
    }

    public get next_page(): string | undefined {
        return this._page.next_page;
    }

    public get prev_page(): string | undefined {
        return this._page.prev_page;
    }

    public async next(): Promise<DefaultPagedCollectionRequest<T>> {
        if (this._page.next_page) {
            return await this._client.api(this._page.next_page).get();
        }
        return Promise.resolve(this); // Maybe return null instead?
    }

    public async prev(): Promise<DefaultPagedCollectionRequest<T>> {
        if (this._page.prev_page) {
            return await this._client.api(this._page.prev_page).get();
        }
        return Promise.resolve(this); // Maybe return null instead?
    }
}

export class DefaultRTClient implements RTClient {
    public readonly rt: RootRequestBuilder;

    public readonly tickets: TicketsRequestBuilder;

    public readonly queues: QueuesRequestBuilder;

    constructor(private readonly _connector: Client) {
        this.queues = new DefaultQueuesRequestBuilder(this._connector);
        this.tickets = new DefaultTicketsRequestBuilder(this._connector);
        this.rt = new DefaultRootRequestBuilder(this._connector);
    }
}

export class DefaultRootRequestBuilder implements RootRequestBuilder {
    constructor(private readonly _connector: Client) {}

    public async get(): Promise<RTRoot> {
        //TypedRTRequest<typeof rtSchemaConfig>
        return createSchemaClientRequest(this._connector.api("rt"), rtSchemaConfig).get();
    }
}

class DefaultQueuesRequestBuilder implements QueuesRequestBuilder {
    constructor(private readonly _connector: Client) {}

    public id(queueId: string): QueueRequestBuilder {
        return new DefaultQueueRequestBuilder(this._connector, queueId);
    }

    public get all(): Promise<QueueRef[]> {
        return this.get().then(async (page: DefaultPagedCollectionRequest<QueueRef>): Promise<QueueRef[]> => {
            const items: QueueRef[] = page.items;
            while (page.next_page) {
                const nextPage = await page.next();
                items.push(...nextPage.items);
                page = nextPage;
            }
            return items;
        });
    }

    public async get(): Promise<DefaultPagedCollectionRequest<QueueRef>> {
        const page = new DefaultPagedCollectionRequest(
            this._connector,
            await this._connector.api("queues/all").get<PagedCollection<QueueRef>>()
        );
        return page;
    }
}

class DefaultQueueRequestBuilder implements QueueRequestBuilder {
    public readonly ticketCustomFields: CustomFieldsRequestBuilder = new DefaultCustomFieldsRequestBuilder(
        this._connector,
        `queue/${this._queueId}`
    );

    constructor(private readonly _connector: Client, private _queueId: string) {}

    public id(queueId: string): QueueRequestBuilder {
        this._queueId = queueId;
        return this;
    }

    public async get(): Promise<Queue> {
        return this._connector.api(`queue/${this._queueId}`).get();
    }
}

class DefaultCustomFieldsRequestBuilder implements CustomFieldsRequestBuilder {
    constructor(private readonly _connector: Client, private readonly _path: string = "") {}

    public id(customFieldId: string): CustomFieldRequestBuilder {
        return new DefaultCustomFieldRequestBuilder(this._connector, customFieldId);
    }

    public get all(): Promise<CustomFieldRef[]> {
        return this.get().then(
            async (page: DefaultPagedCollectionRequest<CustomFieldRef>): Promise<CustomFieldRef[]> => {
                const items: CustomFieldRef[] = page.items;
                while (page.next_page) {
                    const nextPage = await page.next();
                    items.push(...nextPage.items);
                    page = nextPage;
                }
                return items;
            }
        );
    }

    public async get(): Promise<DefaultPagedCollectionRequest<CustomFieldRef>> {
        return new DefaultPagedCollectionRequest(
            this._connector,
            await this._connector.api(`${this._path}/customfields`).get<PagedCollection<CustomFieldRef>>()
        );
    }
}

class DefaultCustomFieldRequestBuilder implements CustomFieldRequestBuilder {
    public customFieldValues: CustomFieldValuesRequestBuilder = new DefaultCustomFieldValuesRequestBuilder(
        this._connector,
        this._customFieldId
    );

    constructor(private readonly _connector: Client, private _customFieldId: string) {}

    public id(customFieldId: string): CustomFieldRequestBuilder {
        this._customFieldId = customFieldId;
        return this;
    }

    public async get(): Promise<CustomField> {
        return this._connector.api(`customfield/${this._customFieldId}`).get();
    }
}

class DefaultCustomFieldValuesRequestBuilder implements CustomFieldValuesRequestBuilder {
    constructor(private readonly _connector: Client, private readonly _customFieldId: string) {}

    public id(customFieldValueId: string): CustomFieldValueRequestBuilder {
        return new DefaultCustomFieldValueRequestBuilder(this._connector, this._customFieldId, customFieldValueId);
    }

    public get all(): Promise<CustomFieldValueRef[]> {
        return this.get().then(
            async (page: DefaultPagedCollectionRequest<CustomFieldValueRef>): Promise<CustomFieldValueRef[]> => {
                const items: CustomFieldValueRef[] = page.items;
                while (page.next_page) {
                    const nextPage = await page.next();
                    items.push(...nextPage.items);
                    page = nextPage;
                }
                return items;
            }
        );
    }

    public async get(): Promise<DefaultPagedCollectionRequest<QueueRef>> {
        return new DefaultPagedCollectionRequest(
            this._connector,
            await this._connector.api(`customfield/${this._customFieldId}/values`).get<PagedCollection<QueueRef>>()
        );
    }
}

class DefaultCustomFieldValueRequestBuilder implements CustomFieldValueRequestBuilder {
    constructor(
        private readonly _connector: Client,
        private readonly _customFieldId: string,
        private _customFieldValueId: string
    ) {}

    public id(customFieldValueId: string): CustomFieldValueRequestBuilder {
        this._customFieldValueId = customFieldValueId;
        return this;
    }

    public async get(): Promise<CustomFieldValue> {
        return this._connector.api(`customfield/${this._customFieldId}/value/${this._customFieldValueId}`).get();
    }
}

class DefaultTicketsRequestBuilder implements TicketsRequestBuilder {
    constructor(private readonly _connector: Client) {}

    public id(ticketId: string): TicketRequestBuilder {
        return new DefaultTicketRequestBuilder(this._connector, ticketId);
    }

    public get all(): Promise<TicketRef[]> {
        return this.get().then(async (page: DefaultPagedCollectionRequest<TicketRef>): Promise<TicketRef[]> => {
            const items: TicketRef[] = page.items;
            while (page.next_page) {
                const nextPage = await page.next();
                items.push(...nextPage.items);
                page = nextPage;
            }
            return items;
        });
    }

    public async get(): Promise<DefaultPagedCollectionRequest<TicketRef>> {
        return new DefaultPagedCollectionRequest(this._connector, await this._connector.api("tickets").get());
    }
}

export interface TicketOptions {
    id: string;
    subject?: string;
    status?: string;
    description?: string;
    requestor?: string;
    owner?: string;
    customFields?: Record<string, any>;
}

class DefaultTicketRequestBuilder implements TicketRequestBuilder {
    public readonly customFields: CustomFieldsRequestBuilder = new DefaultCustomFieldsRequestBuilder(
        this._client,
        `ticket/${this._ticket.id}`
    );

    constructor(private readonly _client: Client, ticketId: string, private _ticket: Partial<TicketOptions> = {}) {
        this._ticket.id = ticketId;
    }

    public subject(subject: string): TicketRequestBuilder {
        this._ticket.subject = subject;
        return this;
    }

    public status(status: string): TicketRequestBuilder {
        this._ticket.status = status;
        return this;
    }

    public description(description: string): TicketRequestBuilder {
        this._ticket.description = description;
        return this;
    }

    public requestor(email: string): TicketRequestBuilder {
        this._ticket.requestor = email;
        return this;
    }

    public owner(email: string): TicketRequestBuilder {
        this._ticket.owner = email;
        return this;
    }

    public customField(fieldName: string, value: string): TicketRequestBuilder {
        if (!this._ticket.customFields) {
            this._ticket.customFields = {};
        }
        this._ticket.customFields[fieldName] = value;
        return this;
    }

    public ticket(ticket: Partial<TicketOptions>): TicketRequestBuilder {
        this._ticket = { ...this._ticket, ...ticket };
        return this;
    }

    public id(ticketId: string): TicketRequestBuilder {
        this._ticket.id = ticketId;
        return this;
    }

    public async get(): Promise<Ticket> {
        return this._client.api(`ticket/${this._ticket.id}`).get();
    }

    public async create(): Promise<CreateTicket> {
        return this._client.api(`ticket`).post(this._ticket) as Promise<CreateTicket>;
    }

    public async save(): Promise<UpdateTicket> {
        return this._client.api(`ticket/${this._ticket.id}`).put(this._ticket) as Promise<UpdateTicket>;
    }
}

// export class APIClient {
//     private _cookie: string | null = null;

//     private readonly _path: string = `${this._config.apiEndpoint}${this._config.apiBasePath}`;

//     constructor(private readonly _config: BotConfiguration) {}

//     public async login(): Promise<string | null> {
//         // Logs into the API and returns the cookie to be used in subsequent requests

//         console.debug(`endpoint:`, this._config.apiEndpoint);

//         // Fetch the cookie from the API using the supplied credentials
//         const cookie: string | null = await fetch(`${this._config.apiEndpoint}`, {
//             method: HttpMethods.Post,
//             headers: {
//                 [HttpHeaders.ContentType]: HttpContentTypes.FormUrlEncoded,
//             },
//             body: new URLSearchParams({
//                 user: this._config.apiUsername,
//                 pass: this._config.apiPassword,
//                 next: "7a73ae647301ce8bdff23044613b37a3",
//             }),
//         })
//             .then((response: Response): string => {
//                 //   const helper = async function (
//                 //     array: ReadableStream<Uint8Array>
//                 //   ): Promise<string> {
//                 //     const utf8 = new TextDecoder("utf-8");
//                 //     let asString = "";
//                 //     for await (const chunk of array as any) {
//                 //       // Can use for-await starting ES 2018
//                 //       asString += utf8.decode(chunk);
//                 //     }
//                 //     return asString;
//                 //   };
//                 //   const asString = await helper(response.body);
//                 //   for (const header of response.headers.entries()) {
//                 //     console.debug(
//                 //       `[${APIClient.name}][DEBUG] ${this.login.name} header:\n${header}`
//                 //     );
//                 //   }

//                 if (!response.ok) {
//                     // If the response is not OK, throw an error to be caught by the catch block below
//                     throw new Error(
//                         `Failed to login to RT API. Status: ${response.status}, Status Text: ${response.statusText}`
//                     );
//                 }

//                 // Gets the Set-Cookie header from the response
//                 const setCookies = response.headers.getSetCookie();
//                 if (setCookies) {
//                     // If the Set-Cookie header is found, return the first cookie in the header (There can be multiple Set-Cookie headers)
//                     return setCookies[0];
//                 }

//                 // If the Set-Cookie header is not found, throw an error to be caught by the catch block below
//                 throw new Error("No Set-Cookie header found in the response.");
//             })
//             .catch((error: any) => {
//                 console.error(error);
//                 return null;
//             });

//         // Returns the result of the login process to the caller
//         return cookie;
//     }

//     public async get<T>(endpoint: string): Promise<T> {
//         this._cookie = await this.login();
//         if (!this._cookie) {
//             throw new Error("Unable to login to the RT API.");
//         }

//         console.debug(`endpoint:`, endpoint);

//         // Fetch the data from the API as a GET request
//         return fetch(endpoint, {
//             method: HttpMethods.Get,
//             headers: {
//                 Cookie: this._cookie,
//                 Accept: HttpContentTypes.Json,
//             },
//         }).then((response: Response): Promise<T> => {
//             // Return the JSON response from the API
//             return response?.json();
//         });
//     }

//     public async next<T>(page: PagedCollection<T>): Promise<PagedCollection<T> | null> {
//         this._cookie = await this.login();
//         if (!this._cookie) {
//             throw new Error("Unable to login to the RT API.");
//         }

//         if (!page.next_page) {
//             // If the next page is not found, return null
//             return null;
//         }

//         console.debug(`endpoint:`, page.next_page);

//         // Fetch the next page of the collection
//         return this.get<PagedCollection<T>>(page.next_page);
//     }

//     public async queues(): Promise<QueuesPage> {
//         const endpoint = `${this._path}/queues/all`;
//         console.debug(`endpoint:`, endpoint);
//         return this.get<QueuesPage>(endpoint);
//     }

//     public async queue(queue: TypedHyperlinkEntity | string): Promise<Queue> {
//         if (typeof queue === "string") {
//             // If the queue is a string, convert it to a hyperlink entity
//             queue = {
//                 id: queue,
//                 _url: `${this._path}/queue/${queue}`,
//                 type: "queue",
//             };
//         }

//         console.debug(`endpoint:`, queue._url);

//         if (queue.type !== "queue") {
//             // If the queue is not of the expected type, throw an error
//             throw new Error(`The supplied hyperlink of type '${queue.type}' is not of the expected type 'queue'.`);
//         }

//         return this.get<Queue>(`${queue._url}`);
//     }

//     public async createTicket(
//         queue: Queue,
//         subject: string,
//         status: string,
//         timeWorked: string,
//         description: string,
//         requestorEmail: string,
//         ownerEmail: string,
//         customFields: {
//             [key: string]: {
//                 text: string;
//                 value: string;
//             };
//         }
//     ): Promise<Ticket> {
//         console.debug(`queue:`, queue);

//         const endpoint = queue?._hyperlinks.find((v: RefHyperlinkEntity) => v.ref === "create")?._url;

//         if (!endpoint) {
//             // If the `create` hyperlink is not found, throw an error
//             throw new Error(`The supplied queue '${queue.id}' does not have a 'create' hyperlink.`);
//         }

//         console.debug(`endpoint:`, endpoint);

//         this._cookie = await this.login();

//         const customFieldsBody: { [key: string]: string } = {};
//         if (customFields && Object.keys(customFields).length > 0) {
//             // If custom fields are provided, convert them to the expected format
//             for (const [_, value] of Object.entries(customFields)) {
//                 customFieldsBody[value.text] = value.value;
//             }
//         }

//         const ownerUser: User = await this.user(ownerEmail);

//         console.debug(`body:`, {
//             Subject: subject,
//             Status: status,
//             Content: description,
//             CustomFields: customFieldsBody,
//             Requestor: requestorEmail,
//             Owner: ownerUser?.Name,
//             TimeWorked: timeWorked,
//         });

//         // Create the ticket using the supplied queue and subject
//         const createTicket: CreateTicket = await fetch(endpoint, {
//             method: HttpMethods.Post,
//             headers: {
//                 Cookie: this._cookie!, // FIXME: Cookie is not guaranteed to be not null
//                 [HttpHeaders.ContentType]: HttpContentTypes.Json,
//             },
//             body: JSON.stringify({
//                 Subject: subject,
//                 Status: status,
//                 Content: description,
//                 CustomFields: customFieldsBody,
//                 // Creator: creator,
//                 Requestor: requestorEmail,
//                 Owner: ownerUser?.Name,
//             }),
//         }).then((response: Response): Promise<CreateTicket> => {
//             return response?.json();
//         });

//         console.debug(`createTicket response:`, createTicket);

//         const ticket = await this.ticket(createTicket);

//         if (
//             timeWorked &&
//             ((typeof timeWorked === "string" && !isNaN(parseInt(timeWorked, 10))) || typeof timeWorked == "number")
//         ) {
//             console.debug(`timeWorked:`, timeWorked);

//             const selfEndpoint = ticket._hyperlinks.find((v) => v.ref === "self")?._url;
//             if (!selfEndpoint) {
//                 throw new Error(`The supplied ticket '${ticket.id}' does not have a 'self' hyperlink`);
//             }

//             console.debug(`selfEndpoint:`, selfEndpoint);

//             // If timeWorked is provided, update the ticket with the time worked
//             const updateTicket = await this.updateTicket(selfEndpoint, {
//                 TimeWorked: timeWorked,
//             }).catch((error: any) => {
//                 console.error(error);
//             });

//             console.debug(`updateTicket response:`, updateTicket);
//         }

//         return ticket;
//     }

//     public async ticket(ticket: TypedHyperlinkEntity): Promise<Ticket> {
//         console.debug(`endpoint:`, ticket._url);

//         if (ticket.type !== "ticket") {
//             throw new Error(`The supplied hyperlink of type '${ticket.type}' is not of the expected type 'ticket'`);
//         }

//         return await this.get<Ticket>(ticket._url);
//     }

//     public async updateTicket(endpoint: string, body: any): Promise<any> {
//         console.debug(`endpoint:`, endpoint);
//         console.debug(`body:`, body);

//         this._cookie = await this.login();

//         return await fetch(endpoint, {
//             method: HttpMethods.Put,
//             headers: {
//                 Cookie: this._cookie!, // FIXME: Cookie is not guaranteed to be not null
//                 [HttpHeaders.ContentType]: HttpContentTypes.Json,
//             },
//             body: JSON.stringify(body),
//         }).then((response: Response): Promise<any> => {
//             return response?.json();
//         });
//     }

//     public async addTicketComment(
//         _graphClient: GraphClient,
//         ticket: Partial<Ticket>,
//         message: ChatMessage
//     ): Promise<string[]> {
//         this._cookie = await this.login();

//         console.debug(`message.id: '${message.id}'`);

//         const endpoint = ticket._hyperlinks?.find((v) => v.ref === "comment")?._url;
//         if (!endpoint) {
//             // If the `comment` hyperlink is not found, throw an error
//             throw new Error(`The supplied ticket '${ticket.id}' does not have a 'comment' hyperlink`);
//         }

//         console.debug(`endpoint:`, endpoint);

//         // const botToken = await MicrosoftGraphUtils.getAccessToken(this._config);

//         const attachments: any[] = [];
//         if (message.body && (message.attachments?.length ?? 0) > 0) {
//             message.body.content += "<br><br>Attachments:<br>";

//             for (const attachment of message.attachments ?? []) {
//                 message.body.content += `<a href="${attachment.contentUrl}">${attachment.name}</a><br>`;

//                 // const buffer = await fetch(attachment.contentUrl, {
//                 //   method: "GET",
//                 //   headers: {
//                 //     Authorization: `${botToken.tokenType} ${botToken.token}`,
//                 //   },
//                 // }).then((res: Response): Promise<ArrayBuffer> => {
//                 //   return res.arrayBuffer();
//                 // });

//                 // console.debug(
//                 //   `[${APIClient.name}][DEBUG] ${
//                 //     this.addTicketComment.name
//                 //   } buffer:\n${JSON.stringify(buffer, null, 2)}`
//                 // );

//                 // const decoder = new TextDecoder("utf-8");
//                 // const bufferStr = decoder.decode(buffer);
//                 // const encodedFile = Buffer.from(bufferStr, "binary").toString("base64");

//                 // console.debug(
//                 //   `[${APIClient.name}][DEBUG] ${
//                 //     this.addTicketComment.name
//                 //   } encodedFile:\n${JSON.stringify(encodedFile, null, 2)}`
//                 // );

//                 // attachments.push({
//                 //   FileName: attachment.id,
//                 //   FileType: "text/plain",
//                 //   FileContent: encodedFile,
//                 // });
//             }
//         }

//         const createComment: string[] = await fetch(endpoint, {
//             method: HttpMethods.Post,
//             headers: {
//                 Cookie: this._cookie!, // FIXME: Cookie is not guaranteed to be not null
//                 [HttpHeaders.ContentType]: HttpContentTypes.Json,
//             },
//             body: JSON.stringify({
//                 Subject: `Respuesta de ${message.from?.user?.displayName}`,
//                 Content: message.body?.content,
//                 ContentType: HttpContentTypes.Html,
//                 TimeTaken: "0",
//                 Attachments: attachments,
//             }),
//         }).then((response: Response): Promise<string[]> => {
//             return response?.json();
//         });

//         console.debug(`createComment response:`, createComment);

//         return createComment;
//     }

//     public async ticketHistory(ticket: Partial<Ticket>): Promise<TicketTransaction> {
//         const endpoint = ticket._hyperlinks?.find((v) => v.ref === "history")?._url;
//         if (!endpoint) {
//             throw new Error(`The supplied ticket '${ticket.id}' does not have a 'history' hyperlink.`);
//         }

//         console.debug(`history endpoint:`, endpoint);

//         const history: TicketTransaction = await this.get<TicketTransaction>(endpoint);

//         console.debug(`history:`, history);

//         return history;
//     }

//     public async queueCustomFields(queueId: string): Promise<CustomField[]> {
//         if (!queueId) {
//             // If the queue ID is not provided, throw an error
//             throw new Error("Argument 'queueId' is required to fetch custom fields");
//         }

//         const queue: Queue = await this.queue(queueId);

//         const customFields: CustomField[] = [];
//         if (!queue || !queue.TicketCustomFields) {
//             // If the queue is not provided or it does not have 'TicketCustomFields', return an empty array
//             return customFields;
//         }

//         for (const hyperlink of queue.TicketCustomFields) {
//             if (hyperlink.ref === "customfield") {
//                 // If the hyperlink is of type 'customfield', fetch the custom fields
//                 const customField: CustomField = await this.get<CustomField>(hyperlink._url);

//                 if (!customField || !customField.id) {
//                     // If the custom field is not found or does not have an ID, skip it
//                     continue;
//                 }

//                 console.debug(`customField:`, customField);

//                 customFields.push(customField);
//             }
//         }

//         return customFields;
//     }

//     public async customFieldValues(customFieldId: string, value: string): Promise<CustomFieldValue[]> {
//         if (!customFieldId) {
//             // If the custom field ID is not provided, throw an error
//             throw new Error("Argument 'customFieldId' is required to fetch custom field values");
//         }

//         const endpoint = `${this._config.apiEndpoint}${this._config.apiBasePath}/customfield/${customFieldId}/values`;

//         console.debug(`endpoint:`, endpoint);
//         console.debug(`value:`, value);

//         // Fetch the custom field choices
//         return await this.get<PagedCollection<TypedHyperlinkEntity>>(endpoint).then(
//             async (response: PagedCollection<TypedHyperlinkEntity>): Promise<CustomFieldValue[]> => {
//                 // Convert the response to an array of CustomFieldValue objects
//                 if (!response || response.items?.length === 0) {
//                     // If the response is empty, return an empty array
//                     return [];
//                 }

//                 // Map the response to CustomFieldValue objects
//                 const customFieldValues: CustomFieldValue[] = [];
//                 for (const ref of response.items) {
//                     if (ref.type !== "customfieldvalue") {
//                         continue;
//                     }

//                     const customFieldValue: CustomFieldValue = await this.get<CustomFieldValue>(ref._url);
//                     if (!customFieldValue || !customFieldValue.id) {
//                         // If the custom field value is not found or does not have an ID, skip it
//                         continue;
//                     }

//                     if (customFieldValue.Category === value) {
//                         customFieldValues.push(customFieldValue);
//                     }
//                 }

//                 const total: number = response.total;
//                 let seen: number = response.count;
//                 let page: number = response.page;
//                 while (seen <= total) {
//                     const nextResponse: PagedCollection<TypedHyperlinkEntity> = await this.get<
//                         PagedCollection<TypedHyperlinkEntity>
//                     >(
//                         `${this._config.apiEndpoint}${
//                             this._config.apiBasePath
//                         }/customfield/${customFieldId}/values?page=${++page}`
//                     );
//                     if (!nextResponse || nextResponse.items?.length === 0) {
//                         break;
//                     }

//                     seen += nextResponse.count;

//                     console.debug(`seen: ${seen}`);

//                     for (const ref of nextResponse.items) {
//                         if (ref.type !== "customfieldvalue") {
//                             continue;
//                         }

//                         const customFieldValue: CustomFieldValue = await this.get<CustomFieldValue>(ref._url);
//                         if (!customFieldValue || !customFieldValue.id) {
//                             // If the custom field value is not found or does not have an id, skip it
//                             continue;
//                         }

//                         if (customFieldValue.Category === value) {
//                             customFieldValues.push(customFieldValue);
//                         }
//                     }
//                 }

//                 return customFieldValues;
//             }
//         );
//     }

//     public async user(email: string): Promise<User> {
//         this._cookie = await this.login();
//         if (!this._cookie) {
//             throw new Error("Unable to login to the RT API.");
//         }

//         console.debug(`email:`, email);

//         // Fetch the user by email
//         return await this.get<User>(`${this._config.apiEndpoint}${this._config.apiBasePath}/user/${email}`);
//     }
// }

const connector = new DefaultClient(
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

export const rt = new DefaultRTClient(connector);

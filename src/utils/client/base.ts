import z, { ZodSchema, ZodType, ZodTypeAny } from "zod";
import { HttpContentTypes, HttpHeaders, HttpMethods } from "../http";

export interface Client {
    api(endpoint: string): ClientRequest;
}

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
        path = path.replace(this._endpoint, "");
        return DefaultClientRequest.create(
            `${this._endpoint}${path.startsWith("/") ? path.slice(1) : path}`,
            this._authProvider
        );
    }
}

export function createClient(
    endpoint: string,
    basePath: string,
    authProvider: () => Promise<{ headerName: string; value: string }>
): Client {
    return new DefaultClient(endpoint, basePath, authProvider);
}

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

    queryParam(name: string, value: QueryParam): this;

    queryParams(params: QueryParams): this;

    header(name: string, value: Header): this;

    headers(headers: Headers): this;

    path(path: string): this;

    body(content: any): this;
}

class DefaultClientRequest implements ClientRequest {
    public static create(
        path: string,
        authProvider: () => Promise<{ headerName: string; value: string }>
    ): ClientRequest {
        return new DefaultClientRequest(authProvider, path);
    }

    private constructor(
        private readonly _authProvider: () => Promise<{ headerName: string; value: string }>,
        private _path: string = "",
        private readonly _queryParams: QueryParams = {},
        private readonly _headers: Headers = {},
        private _body: any = undefined
    ) {
        this.path(this._path ?? "");
    }

    private _url(queryParams: QueryParams): URL {
        const url = new URL(this._path);

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
            body: options?.body ? JSON.stringify(options?.body) : this._body ? JSON.stringify(this._body) : undefined,
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

    public queryParam(name: string, value: QueryParam): this {
        if (value === undefined) {
            delete this._queryParams[name];
            return this;
        }
        this._queryParams[name] = value;
        return this;
    }

    public queryParams(params: QueryParams): this {
        Object.keys(params).forEach((key: string): void => {
            this.queryParam(key, params[key]);
        });
        return this;
    }

    public header(name: string, value: Header): this {
        if (value === undefined) {
            delete this._headers[name];
            return this;
        }
        this._headers[name] = value;
        return this;
    }

    public headers(headers: Headers): this {
        Object.keys(headers).forEach((key: string): void => {
            this.header(key, headers[key]);
        });
        return this;
    }

    public path(path: string): this {
        path = path.trim();
        path = path.startsWith("/") ? path.slice(1) : path;
        if (path.includes("?")) {
            const parts = path.split("?")[0];
            this._path = parts[0];

            if (parts.length > 1) {
                const queryString = parts[1];
                queryString.split("&").forEach((param) => {
                    const [key, value] = param.split("=");
                    this.queryParam(decodeURIComponent(key), decodeURIComponent(value));
                });
            }
        }
        return this;
    }

    public body(content: any): this {
        this._body = content;
        return this;
    }
}

export interface ConfigurableSchemaClientRequest {
    queryParam(name: string, value: QueryParam): this;

    queryParams(params: QueryParams): this;

    header(name: string, value: Header): this;

    headers(headers: Headers): this;
}

export abstract class BaseConfigurableSchemaClientRequest implements ConfigurableSchemaClientRequest {
    constructor(protected readonly request: ClientRequest) {}

    public queryParam(name: string, value: QueryParam): this {
        this.request.queryParam(name, value);
        return this;
    }

    public queryParams(params: QueryParams): this {
        this.request.queryParams(params);
        return this;
    }

    public header(name: string, value: Header): this {
        this.request.header(name, value);
        return this;
    }

    public headers(headers: Headers): this {
        this.request.headers(headers);
        return this;
    }
}

/**
 * Base interface for schema configurations used in schema-based client requests.
 *
 * @public
 */
export interface SchemaConfig {
    /**
     * The path of the resource relative to the base URL.
     */
    path?: string;
}

/**
 * Configuration interface for schema-based client requests, defining Zod schemas for request and response bodies for various HTTP methods.
 *
 * Each property corresponds to a specific HTTP method and endpoint, and indicates the expected schema for that method's request or response body,
 * these schemas are used to validate and type the data sent and received by the client requests. If a property is omitted, it indicates that the
 * corresponding HTTP method is not supported for that resource.
 *
 * @example
 * ```typescript
 * const userSchema = z.object({
 *   id: z.string().uuid(),
 *   name: z.string().min(1),
 *   email: z.string().email(),
 * });
 *
 * const userRequestConfig = {
 *   path: "/user",
 *   getResponse: userSchema,
 *   postRequest: userSchema.omit({ id: true }),
 *   postResponse: userSchema,
 * };
 * ```
 *
 * @public
 */
export interface SchemaClientRequestConfig extends SchemaConfig {
    getResponse?: ZodTypeAny;

    postRequest?: ZodTypeAny;

    postResponse?: ZodTypeAny;

    putRequest?: ZodTypeAny;

    putResponse?: ZodTypeAny;

    deleteResponse?: ZodTypeAny;
}

/**
 * Utility type that infers the type from a Zod schema or returns `undefined` if the schema is not provided.
 */
type SchemaOrUndefined<T> = T extends ZodType<any> ? z.infer<T> : undefined;

/**
 * Utility type that infers the types for various HTTP methods' request and response bodies from a {@link SchemaClientRequestConfig} object.
 * This type maps each property of the configuration to its corresponding inferred type, or `undefined` if the schema is not provided.
 *
 * @example
 * ```typescript
 * const userSchema = z.object({
 *  id: z.string().uuid(),
 *  name: z.string().min(1),
 * email: z.string().email(),
 * });
 *
 * const userRequestConfig: ClientRequestSchemaConfig = {
 *  getResponse: userSchema,
 *  postRequest: userSchema.omit({ id: true }),
 *  postResponse: userSchema,
 * };
 *
 * // The inferred types would be:
 * type InferredTypes = InferFromConfig<typeof userRequestConfig>;
 * // {
 * //   GetResponse: { id: string; name: string; email: string; };
 * //   PostRequest: { name: string; email: string; };
 * //   PostResponse: { id: string; name: string; email: string; };
 * //   PutRequest: undefined;
 * //   PutResponse: undefined;
 * //   DeleteResponse: undefined;
 * // }
 * ```
 *
 * @public
 */
export type InferFromConfig<C extends SchemaClientRequestConfig> = {
    /**
     * The schema for the `GET` response.
     */
    GetResponse: SchemaOrUndefined<C["getResponse"]>;

    /**
     * The schema for the `POST` request.
     */
    PostRequest: SchemaOrUndefined<C["postRequest"]>;

    /**
     * The schema for the `POST` response.
     */
    PostResponse: SchemaOrUndefined<C["postResponse"]>;

    /**
     * The schema for the `PUT` request.
     */
    PutRequest: SchemaOrUndefined<C["putRequest"]>;

    /**
     * The schema for the `PUT` response.
     */
    PutResponse: SchemaOrUndefined<C["putResponse"]>;

    /**
     * The schema for the `DELETE` response.
     */
    DeleteResponse: SchemaOrUndefined<C["deleteResponse"]>;
};

export type SchemaClientRequest<Config extends SchemaClientRequestConfig> = ConfigurableSchemaClientRequest &
    (InferFromConfig<Config>["GetResponse"] extends undefined
        ? {}
        : {
              get: () => Promise<InferFromConfig<Config>["GetResponse"]>;
          }) &
    (InferFromConfig<Config>["PostResponse"] extends undefined
        ? {}
        : {
              post: (
                  content: InferFromConfig<Config>["PostRequest"]
              ) => Promise<InferFromConfig<Config>["PostResponse"]>;
          }) &
    (InferFromConfig<Config>["PutResponse"] extends undefined
        ? {}
        : {
              put: (content: InferFromConfig<Config>["PutRequest"]) => Promise<InferFromConfig<Config>["PutResponse"]>;
          }) &
    (InferFromConfig<Config>["DeleteResponse"] extends undefined
        ? {}
        : {
              delete: () => Promise<InferFromConfig<Config>["DeleteResponse"]>;
          });

export type BeforeCallbacks<Config extends SchemaClientRequestConfig> = {
    get?: () => Promise<void>;

    post?: (content: InferFromConfig<Config>["PostRequest"]) => Promise<InferFromConfig<Config>["PostRequest"]>;

    put?: (content: InferFromConfig<Config>["PutRequest"]) => Promise<InferFromConfig<Config>["PutRequest"]>;

    delete?: () => Promise<void>;
};

export type AfterCallbacks<Config extends SchemaClientRequestConfig> = {
    get?: (response: unknown) => Promise<InferFromConfig<Config>["GetResponse"]>;

    post?: (response: unknown) => Promise<InferFromConfig<Config>["PostResponse"]>;

    put?: (response: unknown) => Promise<InferFromConfig<Config>["PutResponse"]>;

    delete?: (response: unknown) => Promise<InferFromConfig<Config>["DeleteResponse"]>;
};

export type Callbacks<C extends SchemaClientRequestConfig> = {
    before?: BeforeCallbacks<C>;

    after?: AfterCallbacks<C>;
};

class DefaultSchemaClientRequest<C extends SchemaClientRequestConfig> extends BaseConfigurableSchemaClientRequest {
    public static create<C extends SchemaClientRequestConfig>(
        request: ClientRequest,
        config: C,
        callbacks?: Callbacks<C>
    ): SchemaClientRequest<C> {
        return new DefaultSchemaClientRequest<C>(request, config, callbacks);
    }

    private constructor(
        request: ClientRequest,
        private readonly _config: C,
        private readonly _callbacks?: Callbacks<C>
    ) {
        super(request);
    }

    public async get(): Promise<InferFromConfig<C>["GetResponse"]> {
        if (!this._config.getResponse) {
            throw new Error("Resource does not support 'GET'");
        }
        await this._callbacks?.before?.get?.();
        let response = await this.request.get();
        response = this._callbacks?.after?.get ? await this._callbacks.after.get(response) : response;
        return this._config.getResponse.parse(response);
    }

    public async post(content: InferFromConfig<C>["PostRequest"]): Promise<InferFromConfig<C>["PostResponse"]> {
        if (!this._config.postRequest || !this._config.postResponse) {
            throw new Error("Resource does not support 'POST'");
        }
        content = this._callbacks?.before?.post ? await this._callbacks.before.post(content) : content;
        const validatedContent = this._config.postRequest?.parse(content);
        let response = await this.request.post(validatedContent);
        response = this._callbacks?.after?.post ? await this._callbacks.after.post(response) : response;
        return this._config.postResponse.parse(response);
    }

    public async put(content: InferFromConfig<C>["PutRequest"]): Promise<InferFromConfig<C>["PutResponse"]> {
        if (!this._config.putRequest || !this._config.putResponse) {
            throw new Error("Resource does not support 'PUT'");
        }
        content = this._callbacks?.before?.put ? await this._callbacks.before.put(content) : content;
        const validatedContent = this._config.putRequest?.parse(content);
        let response = await this.request.put(validatedContent);
        response = this._callbacks?.after?.put ? await this._callbacks.after.put(response) : response;
        return this._config.putResponse.parse(response);
    }

    public async delete(): Promise<InferFromConfig<C>["DeleteResponse"]> {
        if (!this._config.deleteResponse) {
            throw new Error("Resource does not support 'DELETE'");
        }
        await this._callbacks?.before?.delete?.();
        let response = await this.request.delete();
        response = this._callbacks?.after?.delete ? await this._callbacks.after.delete(response) : response;
        return this._config.deleteResponse.parse(response);
    }
}

export function createSchemaClientRequest<C extends SchemaClientRequestConfig>(
    request: ClientRequest,
    config: C,
    callbacks?: Callbacks<C>
): SchemaClientRequest<C> {
    if (!request) {
        throw new Error("Argument 'request' must be a valid 'ClientRequest' instance.");
    }
    if (!config) {
        throw new Error("Argument 'config' must be a valid 'SchemaConfig' instance.");
    }

    const base = DefaultSchemaClientRequest.create<C>(request, config, callbacks);
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

export interface PagedCollection<_T> {
    // Intentionally left empty
}

export interface PagedSchemaClientRequestConfig<T extends PagedCollection<InferItemFromCollection<T>>>
    extends SchemaConfig {
    getResponse?: ZodSchema<T>;
}

export type InferItemFromConfig<C extends PagedSchemaClientRequestConfig<any>> =
    C extends PagedSchemaClientRequestConfig<infer _P extends PagedCollection<infer T>> ? T : never;

export type InferItemFromCollection<P extends PagedCollection<any>> = P extends PagedCollection<infer T> ? T : never;

export type InferCollectionFromConfig<C extends PagedSchemaClientRequestConfig<any>> =
    C extends PagedSchemaClientRequestConfig<infer P extends PagedCollection<any>> ? P : never;

export type PagedSchemaClientRequest<C extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<C>>> =
    (InferFromConfig<C>["GetResponse"] extends undefined
        ? {}
        : {
              get: () => Promise<InferCollectionFromConfig<C>>;
          }) &
        ConfigurableSchemaClientRequest;

// class DefaultPagedSchemaClientRequest<
//     C extends PagedClientRequestSchemaConfig<InferCollectionFromConfig<C>>
// > extends BaseConfigurableSchemaClientRequest {
//     public static create<C extends PagedClientRequestSchemaConfig<InferCollectionFromConfig<C>>>(
//         request: ClientRequest,
//         config: C,
//         callback?: (response: unknown) => Promise<InferCollectionFromConfig<C>>
//     ): DefaultPagedSchemaClientRequest<C> {
//         return new DefaultPagedSchemaClientRequest(request, config, callback);
//     }

//     private constructor(
//         request: ClientRequest,
//         private readonly _config: C,
//         private readonly _callback?: (response: unknown) => Promise<InferCollectionFromConfig<C>>
//     ) {
//         super(request);
//     }

//     public async get(): Promise<InferCollectionFromConfig<C>> {
//         if (!this._config.getResponse) {
//             throw new Error("Resource does not support 'GET'");
//         }
//         let response = await this.request.get();
//         response = this._callback ? await this._callback(response) : response;
//         return this._config.getResponse.parse(response);
//     }
// }

// export function createPagedSchemaClientRequest<C extends PagedClientRequestSchemaConfig<InferCollectionFromConfig<C>>>(
//     request: ClientRequest,
//     config: C,
//     callback?: (response: unknown) => Promise<InferCollectionFromConfig<C>>
// ): PagedSchemaClientRequest<C> {
//     if (!request) {
//         throw new Error("Argument 'request' must be a valid 'ClientRequest' instance.");
//     }
//     if (!config) {
//         throw new Error("Argument 'config' must be a valid 'SchemaConfig' instance.");
//     }

//     const base = DefaultPagedSchemaClientRequest.create<C>(request, config, callback);
//     const isEnabled: Record<"get", boolean> = {
//         get: !!config.getResponse,
//     };

//     // Wrap in Proxy to intercept access to disabled methods
//     return new Proxy(base, {
//         get(target: DefaultPagedSchemaClientRequest<C>, propertyName: string | symbol, receiver: any): any {
//             if (propertyName in base) {
//                 let keyName = propertyName as keyof typeof isEnabled;
//                 if (!isEnabled[keyName]) {
//                     throw new Error(`Attempting to access disabled property/method '${String(propertyName)}'`);
//                 }
//                 const baseKeyName = keyName as typeof keyName & keyof typeof base;
//                 const value = base[baseKeyName];
//                 if (typeof value === "function") {
//                     return value.bind(base);
//                 }
//                 return value;
//             }

//             const value = Reflect.get(target, propertyName, receiver);
//             if (typeof value === "function") {
//                 return value.bind(target);
//             }
//             return value;
//         },
//     }) as PagedSchemaClientRequest<C>;
// }

export interface SchemaClientRequestBuilder<C extends SchemaClientRequestConfig> {
    request: SchemaClientRequest<C>;
}

export interface PagedSchemaClientRequestBuilder<C extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<C>>>
    extends SchemaClientRequestBuilder<C> {
    request: PagedSchemaClientRequest<C>;
}

export abstract class BaseSchemaClientRequestBuilder<C extends SchemaClientRequestConfig>
    implements SchemaClientRequestBuilder<C>
{
    constructor(
        protected readonly client: Client,
        protected readonly config: C,
        private readonly _callbacks?: Callbacks<C>,
        private readonly _variables: Record<string, string | number | boolean> = {}
    ) {
        if (!client) {
            throw new Error("Argument 'client' must be a valid 'Client' instance.");
        }
        if (!config) {
            throw new Error("Argument 'config' must be a valid 'SchemaConfig' instance.");
        }
    }

    protected variable(name: string, value: string): this {
        this._variables[name] = value;
        return this;
    }

    protected variables(vars: Record<string, string>): this {
        Object.keys(vars).forEach((key: string): void => {
            this.variable(key, vars[key]);
        });
        return this;
    }

    protected path(): string {
        let path = this.config.path;
        // if (typeof path === "function") {
        //     path = path(this);
        // } else
        if (typeof path === "string") {
            const pathStr: string = path;
            Object.entries(this._variables).forEach(([key, value]: [string, string | number | boolean]): void => {
                path = pathStr.replace(`{${key}}`, encodeURIComponent(String(value)));
            });
        } else {
            throw new Error("Cannot build path: 'path' is not defined in the configuration.");
        }
        return path;
    }

    public get request(): SchemaClientRequest<C> {
        let path = this.path();
        if (!path) {
            throw new Error("Cannot build request: computed path is invalid.");
        }
        return createSchemaClientRequest<C>(this.client.api(path), this.config, this._callbacks);
    }
}

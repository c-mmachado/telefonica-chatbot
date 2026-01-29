import { ZodType, ZodTypeAny } from "zod";
import { HttpContentTypes, HttpHeaders, HttpMethod } from "../http";

export interface Client {
    api(endpoint: string): ClientRequest;
}

class DefaultClient implements Client {
    private readonly _endpoint: string;

    constructor(
        endpoint: string,
        basePath: string,
        private readonly _authProvider?: () => Promise<{ headerName: string; value: string }>,
    ) {
        this._endpoint = `${endpoint}${endpoint.endsWith("/") ? "" : "/"}${
            basePath.startsWith("/") ? basePath.slice(1) : basePath
        }${basePath.endsWith("/") ? "" : "/"}`;
    }

    public api(path: string): ClientRequest {
        path = path.replace(this._endpoint, "");
        return DefaultClientRequest.create(
            `${this._endpoint}${path.startsWith("/") ? path.slice(1) : path}`,
            this._authProvider,
        );
    }
}

export function createClient(
    endpoint: string,
    basePath: string,
    authProvider?: () => Promise<{ headerName: string; value: string }>,
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
    get<GetResponse>(content?: unknown): Promise<GetResponse>;

    post<PostResponse>(): Promise<PostResponse>;
    post<PostResponse>(content?: unknown): Promise<PostResponse>;

    put<PutResponse>(): Promise<PutResponse>;
    put<PutResponse>(content?: unknown): Promise<PutResponse>;

    delete<DeleteResponse>(): Promise<DeleteResponse>;
    delete<DeleteResponse>(content?: unknown): Promise<DeleteResponse>;

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
        authProvider?: () => Promise<{ headerName: string; value: string }>,
    ): ClientRequest {
        return new DefaultClientRequest(authProvider, path);
    }

    private constructor(
        private readonly _authProvider?: () => Promise<{ headerName: string; value: string }>,
        private _path: string = "",
        private readonly _queryParams: QueryParams = {},
        private readonly _headers: Headers = {},
        private _body: any = undefined,
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
        method: HttpMethod,
        options?: { body?: unknown; headers?: Headers; queryParams?: QueryParams },
    ): Promise<any> {
        let headers = {
            Accept: HttpContentTypes.Json,
            ...this._toHeaders(options?.headers ?? {}),
        };

        const auth = await this._authProvider?.();
        if (auth && auth.headerName && auth.value) {
            console.debug("Adding authentication header to request:", auth);
            headers = {
                ...headers,
                [auth.headerName]: auth.value,
            };
        }

        console.debug(`Making request to resource '${method?.toUpperCase()} ${this._path}'`);
        console.debug(`options:`, options);

        return fetch(this._url(options?.queryParams ?? {}), {
            method: method.toUpperCase(),
            headers: headers,
            body: options?.body ? JSON.stringify(options?.body) : this._body ? JSON.stringify(this._body) : undefined,
        })
            .then((response: Response): Promise<any> => {
                if (!response.ok) {
                    throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
                }
                return response.json();
            })
            .catch((error: any): Promise<any> => {
                console.error(`An error occurred during a request to resource '${method} ${this._path}'`);
                return Promise.reject(error);
            });
    }

    public async get<GetResponse>(): Promise<GetResponse>;
    public async get<GetResponse>(content?: unknown): Promise<GetResponse> {
        return this._request(HttpMethod.Get, {
            body: content,
            headers: {
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public async post<PostResponse>(): Promise<PostResponse>;
    public async post<PostResponse>(content?: unknown): Promise<PostResponse> {
        return this._request(HttpMethod.Post, {
            body: content,
            headers: {
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public async put<PutResponse>(): Promise<PutResponse>;
    public async put<PutResponse>(content?: unknown): Promise<PutResponse> {
        return this._request(HttpMethod.Put, {
            body: content,
            headers: {
                [HttpHeaders.ContentType]: HttpContentTypes.Json,
                ...this._headers,
            },
            queryParams: this._queryParams,
        });
    }

    public async delete<DeleteResponse>(): Promise<DeleteResponse>;
    public async delete<DeleteResponse>(content?: unknown): Promise<DeleteResponse> {
        return this._request(HttpMethod.Delete, {
            body: content,
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
            const parts = path.split("?");
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
 * Base interface for endpoint configurations used in schema-based client requests.
 *
 * @public
 */
export interface EndpointConfig {
    /**
     * The path of the resource relative to the base URL.
     */
    path?: string;
}

type ResponseMethodConfig<Response = any> = {
    response: ZodType<Response>;
};

type BodyMethodConfig<Request = any, Response = any> = {
    request: ZodType<Request>;
    response: ZodType<Response>;
};

export type MethodConfig<Request = any, Response = any> =
    | BodyMethodConfig<Request, Response>
    | ResponseMethodConfig<Response>;

export interface SchemaEndpointConfig<
    Methods extends Partial<Record<HttpMethod, MethodConfig<any, any>>> = {},
> extends EndpointConfig {
    methods?: Methods;
}

// export interface SchemaClientRequestConfig extends SchemaConfig {
//     getResponse?: ZodTypeAny;

//     postRequest?: ZodTypeAny;

//     postResponse?: ZodTypeAny;

//     putRequest?: ZodTypeAny;

//     putResponse?: ZodTypeAny;

//     deleteResponse?: ZodTypeAny;
// }

export function createSchemaEndpointConfig<
    const Config extends SchemaEndpointConfig<Methods>,
    Methods extends Partial<Record<HttpMethod, MethodConfig<any, any>>> = {},
>(config: Config): Config {
    return config;
}

/**
 * Utility type that infers the type from a Zod schema or returns `undefined` if the schema is not provided.
 */
// type SchemaOrUndefined<T> = T extends ZodType<any> ? z.infer<T> : undefined;

// export type InferFromConfig<C extends SchemaClientRequestConfig> = {
//     /**
//      * The schema for the `GET` response.
//      */
//     GetResponse: SchemaOrUndefined<C["getResponse"]>;

//     /**
//      * The schema for the `POST` request.
//      */
//     PostRequest: SchemaOrUndefined<C["postRequest"]>;

//     /**
//      * The schema for the `POST` response.
//      */
//     PostResponse: SchemaOrUndefined<C["postResponse"]>;

//     /**
//      * The schema for the `PUT` request.
//      */
//     PutRequest: SchemaOrUndefined<C["putRequest"]>;

//     /**
//      * The schema for the `PUT` response.
//      */
//     PutResponse: SchemaOrUndefined<C["putResponse"]>;

//     /**
//      * The schema for the `DELETE` response.
//      */
//     DeleteResponse: SchemaOrUndefined<C["deleteResponse"]>;
// };

export type InferFromConfig<Config extends SchemaEndpointConfig> = Config extends {
    methods: infer Methods extends Partial<Record<HttpMethod, MethodConfig<any, any>>>;
}
    ? {
          [K in keyof Methods]: {
              Request: Methods[K] extends BodyMethodConfig<infer Request, any> ? Request : undefined;
              Response: Methods[K] extends BodyMethodConfig<any, infer Response> | ResponseMethodConfig<infer Response>
                  ? Response
                  : undefined;
          };
      }
    : {};

// export type SchemaClientRequest<Config extends SchemaClientRequestConfig> = ConfigurableSchemaClientRequest &
//     (InferFromConfig<Config>["GetResponse"] extends undefined
//         ? {}
//         : {
//               get: () => Promise<InferFromConfig<Config>["GetResponse"]>;
//           }) &
//     (InferFromConfig<Config>["PostResponse"] extends undefined
//         ? {}
//         : {
//               post: (
//                   content: InferFromConfig<Config>["PostRequest"]
//               ) => Promise<InferFromConfig<Config>["PostResponse"]>;
//           }) &
//     (InferFromConfig<Config>["PutResponse"] extends undefined
//         ? {}
//         : {
//               put: (content: InferFromConfig<Config>["PutRequest"]) => Promise<InferFromConfig<Config>["PutResponse"]>;
//           }) &
//     (InferFromConfig<Config>["DeleteResponse"] extends undefined
//         ? {}
//         : {
//               delete: () => Promise<InferFromConfig<Config>["DeleteResponse"]>;
//           });

export type SchemaClientRequest<Config extends SchemaEndpointConfig> = ConfigurableSchemaClientRequest & {
    [Method in keyof InferFromConfig<Config>]: InferFromConfig<Config>[Method] extends {
        Request: infer Request;
        Response: infer Response;
    }
        ? Request extends undefined
            ? Response extends undefined
                ? never
                : () => Promise<Response>
            : Response extends undefined
              ? (content: Request) => Promise<void>
              : (content: Request) => Promise<Response>
        : never;
};

type MethodType<Config extends SchemaEndpointConfig, Method extends HttpMethod> =
    InferFromConfig<Config> extends {
        [Method in HttpMethod]?: {
            Request: any;
            Response: any;
        };
    }
        ? InferFromConfig<Config>[Method]
        : never;

export type MethodRequestType<Config extends SchemaEndpointConfig, Method extends HttpMethod> =
    MethodType<Config, Method> extends { Request: infer Request } ? Request : never;

export type MethodResponseType<Config extends SchemaEndpointConfig, Method extends HttpMethod> =
    MethodType<Config, Method> extends { Response: infer Response } ? Response : never;

export type BeforeCallbacks<Config extends SchemaEndpointConfig> = Partial<{
    [Method in keyof InferFromConfig<Config>]: InferFromConfig<Config>[Method] extends {
        Request: infer Request;
    }
        ? Request extends undefined
            ? () => Promise<void>
            : (content: Request) => Promise<Request>
        : never;
}>;

export type AfterCallbacks<Config extends SchemaEndpointConfig> = Partial<{
    [Method in keyof InferFromConfig<Config>]?: InferFromConfig<Config>[Method] extends {
        Response: infer Response;
    }
        ? Response extends undefined
            ? (response: unknown) => Promise<void>
            : (response: unknown) => Promise<Response>
        : never;
}>;

export type Callbacks<Config extends SchemaEndpointConfig> = {
    before?: BeforeCallbacks<Config>;

    after?: AfterCallbacks<Config>;
};

class DefaultSchemaClientRequest<Config extends SchemaEndpointConfig> extends BaseConfigurableSchemaClientRequest {
    public static create<Config extends SchemaEndpointConfig>(
        request: ClientRequest,
        config: Config,
        callbacks?: Callbacks<Config>,
    ): SchemaClientRequest<Config> {
        const base = new DefaultSchemaClientRequest<Config>(request, config, callbacks);

        return new Proxy(base, {
            get(target: DefaultSchemaClientRequest<Config>, propertyName: string | symbol, receiver: any): any {
                if (
                    typeof propertyName === "string" &&
                    (Object.values(HttpMethod).includes(propertyName as HttpMethod) || propertyName in HttpMethod) &&
                    base._supportsMethod(propertyName as HttpMethod)
                ) {
                    // const keyName = propertyName as keyof HttpMethod;
                    // return base.method.bind(base);

                    const method: HttpMethod | undefined = Object.entries(HttpMethod).find(
                        (value: [string, HttpMethod], _index: number, _array: [string, HttpMethod][]) => {
                            return value[0] === propertyName || value[1] === propertyName;
                        },
                    )?.[1];

                    return (content: any) => {
                        return base.method(content, method as HttpMethod & keyof InferFromConfig<Config>);
                    };
                }

                const value = Reflect.get(target, propertyName, receiver);

                if (typeof value !== "function") {
                    return target[propertyName as keyof typeof target];
                }

                const boundFunction: Function = value.bind(target);
                return (...args: any) => {
                    try {
                        const result = boundFunction.apply(target, args);
                        if (result == target) {
                            return receiver;
                        }
                        return result;
                    } catch (e: any) {
                        throw e;
                    }
                };
            },
        }) as SchemaClientRequest<Config>;
    }

    private constructor(
        request: ClientRequest,
        private readonly _config: Config,
        private readonly _callbacks?: Callbacks<Config>,
    ) {
        super(request);
    }

    private _supportsMethod(method: HttpMethod): boolean {
        return !!this._config?.methods && method in this._config.methods;
    }

    private _schemaForMethod(method: HttpMethod, schema: "request" | "response"): ZodTypeAny | undefined {
        if (!this._supportsMethod(method)) {
            return undefined;
        }

        const methods: {
            [method]: Partial<Record<"request" | "response", ZodTypeAny>>;
        } = this._config.methods ?? {};

        if (schema in methods[method]) {
            return methods[method][schema] as ZodType;
        }

        return undefined;
    }

    public async method<Method extends HttpMethod & keyof InferFromConfig<Config>>(
        content: MethodRequestType<Config, Method>,
        method: Method,
    ): Promise<MethodResponseType<Config, Method>> {
        if (!this._supportsMethod(method)) {
            throw new Error(`Resource does not support method '${method}'`);
        }

        const requestSchema = this._schemaForMethod(method, "request");
        const hasRequestSchema = !!requestSchema && requestSchema instanceof ZodType;
        if (!hasRequestSchema && content) {
            console.warn(
                `No request schema defined for method '${method}', but content was provided. It will be ignored.`,
            );
        }

        content = (
            this._callbacks?.before?.[method] ? await this._callbacks.before[method](content) : content
        ) as MethodRequestType<Config, Method>;
        const validatedContent = hasRequestSchema ? requestSchema.parse(content) : undefined;

        if (!(method in this.request)) {
            throw new Error(`Method '${method}' is not implemented in the underlying 'ClientRequest'.`);
        }
        const requestKey = method as keyof ClientRequest;
        if (typeof this.request[requestKey] !== "function") {
            throw new Error(`Method '${method}' is not a function in the underlying 'ClientRequest'.`);
        }
        const requestMethod = this.request[requestKey] as Function;

        let response: unknown;
        // if (this._config.path === "/queues/all") {
        //     response = tmp;
        // } else if (this._config.path === "/queue/{id}") {
        //     response = tmp2;
        // } else {
        // }

        if (hasRequestSchema) {
            response = await requestMethod.call(this.request, validatedContent);
        } else {
            response = await requestMethod.call(this.request);
        }
        response = this._callbacks?.after?.[method] ? await this._callbacks.after[method](response) : response;

        const responseSchema = this._schemaForMethod(method, "response");
        const hasResponseSchema = !!responseSchema && responseSchema instanceof ZodType;
        if (!hasResponseSchema && response) {
            console.warn(
                `No response schema defined for method '${method}', but response was received. It will be ignored.`,
            );
        }
        if (!hasResponseSchema) {
            return undefined as MethodResponseType<Config, Method>;
        }

        // responseSchema.parse(response); // TODO: Reenable validation
        return response as MethodResponseType<Config, Method>;
    }

    // public async get(): Promise<MethodResponseType<Config, HttpMethod.Get>> {
    //     if (!this._supportsMethod(HttpMethod.Get)) {
    //         throw new Error(`Resource does not support method '${HttpMethod.Get}'`);
    //     }
    //     await this._callbacks?.before?.get?.();
    //     let response = await this.request.get();
    //     response = this._callbacks?.after?.get ? await this._callbacks.after.get(response) : response;
    //     return this._schemaForMethod(HttpMethod.Get, "response")?.parse(response);
    // }

    // public async post(
    //     content: MethodRequestType<Config, HttpMethod.Post>
    // ): Promise<MethodResponseType<Config, HttpMethod.Post>> {
    //     if (!this._supportsMethod(HttpMethod.Post)) {
    //         throw new Error(`Resource does not support method '${HttpMethod.Post}'`);
    //     }
    //     content = this._callbacks?.before?.post ? await this._callbacks.before.post(content) : content;
    //     const validatedContent = this._schemaForMethod(HttpMethod.Post, "request")?.parse(content);
    //     let response = await this.request.post(validatedContent);
    //     response = this._callbacks?.after?.post ? await this._callbacks.after.post(response) : response;
    //     return this._schemaForMethod(HttpMethod.Post, "response")?.parse(response);
    // }

    // public async put(
    //     content: MethodRequestType<Config, HttpMethod.Put>
    // ): Promise<MethodResponseType<Config, HttpMethod.Put>> {
    //     if (!this._supportsMethod(HttpMethod.Put)) {
    //         throw new Error(`Resource does not support method '${HttpMethod.Put}'`);
    //     }
    //     content = this._callbacks?.before?.put ? await this._callbacks.before.put(content) : content;
    //     const validatedContent = this._schemaForMethod(HttpMethod.Put, "request")?.parse(content);
    //     let response = await this.request.put(validatedContent);
    //     response = this._callbacks?.after?.put ? await this._callbacks.after.put(response) : response;
    //     return this._schemaForMethod(HttpMethod.Put, "response")?.parse(response);
    // }

    // public async delete(): Promise<MethodResponseType<Config, HttpMethod.Delete>> {
    //     if (!this._supportsMethod(HttpMethod.Delete)) {
    //         throw new Error(`Resource does not support method '${HttpMethod.Delete}'`);
    //     }
    //     await this._callbacks?.before?.delete?.();
    //     let response = await this.request.delete();
    //     response = this._callbacks?.after?.delete ? await this._callbacks.after.delete(response) : response;
    //     return this._schemaForMethod(HttpMethod.Delete, "response")?.parse(response);
    // }
}

function createSchemaClientRequest<Config extends SchemaEndpointConfig>(
    request: ClientRequest,
    config: Config,
    callbacks?: Callbacks<Config>,
): SchemaClientRequest<Config> {
    if (!request) {
        throw new Error("Argument 'request' must be a valid 'ClientRequest' instance.");
    }
    if (!config) {
        throw new Error("Argument 'config' must be a valid 'SchemaConfig' instance.");
    }

    // const base =
    return DefaultSchemaClientRequest.create<Config>(request, config, callbacks);
    // const isEnabled: Record<"get" | "post" | "put" | "delete", boolean> = {
    //     get: !!config.getResponse,
    //     post: !!(config.postRequest && config.postResponse),
    //     put: !!(config.putRequest && config.putResponse),
    //     delete: !!config.deleteResponse,
    // };

    // Wrap in Proxy to intercept access to disabled methods
    // return new Proxy(base, {
    //     get(target: DefaultSchemaClientRequest<C>, propertyName: string | symbol, receiver: any): any {
    //         if (propertyName in base) {
    //             let keyName = propertyName as keyof typeof isEnabled;
    //             if (!isEnabled[keyName]) {
    //                 throw new Error(`Attempting to access disabled property/method '${String(propertyName)}'`);
    //             }

    //             const baseKeyName = keyName as typeof keyName & keyof typeof base;
    //             const value = base[baseKeyName];
    //             if (typeof value === "function") {
    //                 return value.bind(base);
    //             }
    //             return value;
    //         }

    //         const value = Reflect.get(target, propertyName, receiver);
    //         if (typeof value === "function") {
    //             return value.bind(target);
    //         }
    //         return value;
    //     },
    // }) as SchemaClientRequest<C>;
}

// export interface PagedCollection<_T> {
//     // Intentionally left empty
// }

// export interface PagedSchemaClientRequestConfig<Collection extends PagedCollection<InferItemFromCollection<Collection>>>
//     extends SchemaConfig {
//     methods: {
//         get: ResponseMethodConfig<Collection>;
//     };
// }

// export type InferItemFromConfig<Config extends PagedSchemaClientRequestConfig<any>> =
//     Config extends PagedSchemaClientRequestConfig<infer _P extends PagedCollection<infer T>> ? T : never;

// export type InferItemFromCollection<Collection extends PagedCollection<any>> = Collection extends PagedCollection<
//     infer Item
// >
//     ? Item
//     : never;

// export type InferCollectionFromConfig<Config extends PagedSchemaClientRequestConfig<any>> =
//     Config extends PagedSchemaClientRequestConfig<infer Collection extends PagedCollection<any>> ? Collection : never;

// export type PagedSchemaClientRequest<Config extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<Config>>> =
//     ConfigurableSchemaClientRequest &
//         (MethodRequestType<Config, HttpMethod.Get> extends undefined
//             ? MethodResponseType<Config, HttpMethod.Get> extends undefined
//                 ? {}
//                 : {
//                       get: () => Promise<MethodResponseType<Config, HttpMethod.Get>>;
//                   }
//             : {
//                   get: (
//                       content: MethodRequestType<Config, HttpMethod.Get>
//                   ) => Promise<MethodResponseType<Config, HttpMethod.Get>>;
//               });

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

export interface SchemaEndpointConfigurer<Config extends SchemaEndpointConfig> {
    request: SchemaClientRequest<Config>;
}

// export interface PagedSchemaClientRequestBuilder<
//     C extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<C>>
// > {
//     request: PagedSchemaClientRequest<C>;
// }

export abstract class BaseSchemaEndpointConfigurer<
    Config extends SchemaEndpointConfig,
> implements SchemaEndpointConfigurer<Config> {
    constructor(
        protected readonly client: Client,
        protected readonly config: Config,
        protected readonly callbacks?: Callbacks<Config>,
        private readonly _variables: Record<string, string | number | boolean> = {},
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
            console.debug("Original path string:", path);
            console.debug("Resolving path with variables:", this._variables);

            Object.entries(this._variables).forEach(([key, value]: [string, string | number | boolean]): void => {
                path = path!.replace(`{${key}}`, encodeURIComponent(String(value)));
            });
        } else {
            throw new Error("Property 'path' must be defined in the endpoint's schema configuration.");
        }
        return path;
    }

    public get request(): SchemaClientRequest<Config> {
        let path = this.path();
        if (!path) {
            throw new Error("Property 'path' could not be resolved for the endpoint.");
        }
        return createSchemaClientRequest<Config>(this.client.api(path), this.config, this.callbacks);
    }
}

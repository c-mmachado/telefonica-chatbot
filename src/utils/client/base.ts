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

        console.info(`Making request to resource '${method?.toUpperCase()} ${this._path}'`);
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

export function createSchemaEndpointConfig<
    const Config extends SchemaEndpointConfig<Methods>,
    Methods extends Partial<Record<HttpMethod, MethodConfig<any, any>>> = {},
>(config: Config): Config {
    return config;
}

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

    return DefaultSchemaClientRequest.create<Config>(request, config, callbacks);
}

export interface SchemaEndpointConfigurer<Config extends SchemaEndpointConfig> {
    request: SchemaClientRequest<Config>;
}

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

    protected variable(name: string, value: string | number | boolean): this {
        this._variables[name] = value;
        return this;
    }

    protected variables(vars: Record<string, string | number | boolean>): this {
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
            // console.debug("Original path string:", path);
            // console.debug("Resolving path with variables:", this._variables);

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

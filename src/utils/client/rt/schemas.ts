import z, { ZodType } from "zod";

import {
    BaseSchemaClientRequestBuilder,
    Client,
    InferCollectionFromConfig,
    InferItemFromConfig,
    PagedSchemaClientRequestConfig,
    PagedCollection,
    PagedSchemaClientRequest,
    PagedSchemaClientRequestBuilder,
} from "../base";

// export type HyperlinkEntity = Partial<{
//     id: string;
//     name: string;

//     type: HyperlinkType;
//     ref: HyperlinkRef;

//     _url: string;

//     from: string;
//     to: string;
//     label: string;
//     update: string;
// }>;

// export type TypedHyperlinkEntity = HyperlinkEntity & Required<HyperlinkEntity, "type" | "_url">;

// export type RefHyperlinkEntity = HyperlinkEntity & Required<HyperlinkEntity, "ref" | "_url">;

// export type CustomFieldHyperlink = TypedHyperlinkEntity &
//     Required<HyperlinkEntity, "name"> & {
//         values?: string[];
//     };

export enum HyperlinkRef {
    Self = "self",
    User = "user",
    Queue = "queue",
    Ticket = "ticket",
    CustomField = "customfield",
    CustomFieldValue = "customfieldvalue",
    Create = "create",
    Comment = "comment",
    History = "history",
}

export enum HyperlinkType {
    User = "user",
    Queue = "queue",
    Ticket = "ticket",
    CustomField = "customfield",
    CustomFieldValue = "customfieldvalue",
    Transaction = "transaction", // ? check
}

// Base types schemas
export const hyperlinkSchema = z.object({
    id: z.string().min(1).optional(),
    name: z.string().min(1).optional(),
    type: z.nativeEnum(HyperlinkType).optional(),
    ref: z.nativeEnum(HyperlinkRef).optional(),
    _url: z.string().url().optional(),
    from: z.string().min(1).optional(),
    to: z.string().min(1).optional(),
    label: z.string().min(1).optional(),
    update: z.string().min(1).optional(),
});
export interface HyperlinkEntity extends z.infer<typeof hyperlinkSchema> {}

export const typedHyperlinkSchema = hyperlinkSchema.and(
    z.object({
        type: z.nativeEnum(HyperlinkType),
        _url: z.string().url(),
    })
);
export interface TypedHyperlinkEntity extends z.infer<typeof typedHyperlinkSchema> {}

export const refHyperlinkSchema = hyperlinkSchema.and(
    z.object({
        ref: z.nativeEnum(HyperlinkRef),
        _url: z.string().url(),
    })
);
export interface RefHyperlinkEntity extends z.infer<typeof refHyperlinkSchema> {}

/**
 * Generic factory function to create paged schemas for RT collections used
 * to allow a strongly typed paged collection response from the RT API.
 *
 * @param itemSchema Schema for the items in the collection
 * @returns Zod schema for the paged collection
 */
export const createPagedCollectionSchema = <T extends z.ZodTypeAny>(itemSchema: T) => {
    return z.object({
        items: z.array(itemSchema),
        page: z.number().min(1),
        per_page: z.number().min(1),
        total: z.number().min(0),
        pages: z.number().min(1),
        count: z.number().min(0),
        next_page: z.string().url().optional(),
        prev_page: z.string().url().optional(),
    });
};

// const configurableSchemaClientRequestSchema = z.object({});

// export const createPagedCollectionSchema = <T extends z.ZodTypeAny>(itemSchema: T) => {
//     // Forward declarations for mutual recursion
//     let pagedSchema: z.ZodObject<any>;
//     let clientRequestSchema: z.ZodType;

//     // Define the client request schema, incorporating the conditional logic.
//     // Based on your type, if "GetResponse" is always defined for paged scenarios, include 'get'.
//     // If it can be undefined, make 'get' optional or handle conditionally (though Zod doesn't support type-level conditionals directly).
//     // Here, assuming it's always present for paged collections.
//     clientRequestSchema = z.lazy(
//         () =>
//             configurableSchemaClientRequestSchema.extend({
//                 get: z.function().returns(z.promise(z.lazy(() => pagedSchema))),
//             })
//         // If GetResponse can be undefined in some configs, you could make 'get' optional:
//         // .extend({ get: z.function().returns(z.promise(z.lazy(() => pagedSchema))).optional() })
//         // But this would approximate the type; adjust based on your exact needs.
//     );

//     pagedSchema = z.object({
//         items: z.array(itemSchema),
//         page: z.number().min(1),
//         per_page: z.number().min(1),
//         total: z.number().min(0),
//         pages: z.number().min(1),
//         count: z.number().min(0),
//         next_page: z.string().url().optional(),
//         prev_page: z.string().url().optional(),

//         next: z
//             .function()
//             .returns(z.promise(z.lazy(() => clientRequestSchema)))
//             .optional(),
//     });

//     return pagedSchema;
// };

// Factory for enhanced schema with recursive 'next' method
export const createNavigatablePagedCollectionSchema = <T extends z.ZodTypeAny>(itemSchema: T) => {
    type Item = z.infer<T>;
    type CollectionOfItem = RTNavigatablePagedCollection<Item>;
    type Config = PagedSchemaClientRequestConfig<CollectionOfItem>;

    // Manual type definitions for inference (since v3 may infer 'any' for recursion)
    type RTNavigatablePagedCollection<Item> = RTPagedCollection<Item> & {
        next: () => PagedSchemaClientRequestBuilder<Config>;

        prev: () => PagedSchemaClientRequestBuilder<Config>;
    };

    // type PagedSchemaClientRequestBuilder<C extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<C>>> =
    //     SchemaClientRequestBuilder<C> & {
    //         request: PagedSchemaClientRequest<C>;
    //     };

    // Base schemas (update with actual definitions if needed)
    const configurableSchemaClientRequestSchema = z.object({
        headers: z
            .function()
            .args(z.record(z.string()))
            .returns(z.lazy((): z.ZodType => configurableSchemaClientRequestSchema)),
        header: z
            .function()
            .args(z.string(), z.string())
            .returns(z.lazy((): z.ZodType => configurableSchemaClientRequestSchema)),
        queryParams: z
            .function()
            .args(z.record(z.string(), z.string()))
            .returns(z.lazy((): z.ZodType => configurableSchemaClientRequestSchema)),
        queryParam: z
            .function()
            .args(z.string(), z.string())
            .returns(z.lazy((): z.ZodType => configurableSchemaClientRequestSchema)),
    });

    // Forward declarations
    let enhancedSchema: z.ZodType<RTNavigatablePagedCollection<Item>>;
    let requestSchema: z.ZodType<PagedSchemaClientRequest<Config>>;
    let builderSchema: z.ZodType<PagedSchemaClientRequestBuilder<Config>>;

    // Assign in dependency order
    enhancedSchema = createPagedCollectionSchema(itemSchema).extend({
        next: z.function().returns(z.lazy(() => builderSchema)),
        prev: z.function().returns(z.lazy(() => builderSchema)),
    });

    requestSchema = configurableSchemaClientRequestSchema.extend({
        get: z.function().returns(z.promise(z.lazy(() => enhancedSchema))),
    });

    builderSchema = z.object({
        request: requestSchema,
    });

    return enhancedSchema;
};

export interface RTPagedCollection<Item>
    extends z.infer<ReturnType<typeof createPagedCollectionSchema<z.ZodType<Item>>>>,
        PagedCollection<Item> {}

export interface RTNavigatablePagedCollection<
    Config extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<Config>>,
    Item extends InferItemFromConfig<Config> = InferItemFromConfig<Config>
> extends RTPagedCollection<Item> {
    items: Item[];

    next(): PagedSchemaClientRequestBuilder<Config>;

    prev(): PagedSchemaClientRequestBuilder<Config>;
}

export function createRTNavigatablePagedCollection<
    Config extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<Config>>,
    Item extends InferItemFromConfig<Config> = InferItemFromConfig<Config>
>(client: Client, config: Config, page: RTPagedCollection<Item>): RTNavigatablePagedCollection<Config, Item> {
    return new DefaultRTNavigatablePagedCollection(client, config, page);
}

class DefaultRTNavigatablePagedCollection<
    Config extends PagedSchemaClientRequestConfig<InferCollectionFromConfig<Config>>,
    Item extends InferItemFromConfig<Config> = InferItemFromConfig<Config>
> implements RTNavigatablePagedCollection<Config, Item>
{
    constructor(
        private readonly _client: Client,
        private readonly _config: Config,
        private readonly _page: RTPagedCollection<Item>
    ) {}

    public get items(): Item[] {
        return this._page.items;
    }

    public get page(): number {
        return this._page.page;
    }

    public get per_page(): number {
        return this._page.per_page;
    }

    public get total(): number {
        return this._page.total;
    }

    public get pages(): number {
        return this._page.pages;
    }

    public get count(): number {
        return this._page.count;
    }

    public get next_page(): string | undefined {
        return this._page.next_page;
    }

    public get prev_page(): string | undefined {
        return this._page.prev_page;
    }

    public next(): PagedSchemaClientRequestBuilder<Config> {
        if (!this._client || !this._config) {
            throw new Error("Cannot get next page: No client or config available.");
        }

        const nextPage = this.next_page;
        class NextEndpointConfigurer
            extends BaseSchemaClientRequestBuilder<Config>
            implements PagedSchemaClientRequestBuilder<Config>
        {
            // TODO: Callbacks should also be propagated to the new instance
            constructor(client: Client, config: Config) {
                super(client, {
                    ...config,
                    path: nextPage,
                });
            }
        }
        return new NextEndpointConfigurer(this._client, this._config);
    }

    public prev(): PagedSchemaClientRequestBuilder<Config> {
        if (!this._client || !this._config) {
            throw new Error("Cannot get previous page: No client or config available.");
        }

        const prevPage = this.prev_page;
        class PrevEndpointConfigurer
            extends BaseSchemaClientRequestBuilder<Config>
            implements PagedSchemaClientRequestBuilder<Config>
        {
            // TODO: Callbacks should also be propagated to the new instance
            constructor(client: Client, config: Config) {
                super(client, {
                    ...config,
                    path: prevPage,
                });
            }
        }
        return new PrevEndpointConfigurer(this._client, this._config);
    }
}

/**
 * /rt endpoint schemas
 */
const rtGetResponseSchema = z.object({
    Version: z.string().min(1),
});
interface RTGetResponse extends z.infer<typeof rtGetResponseSchema> {}
export const rtSchemaConfig = {
    path: "/rt",
    getResponse: rtGetResponseSchema as ZodType<RTGetResponse>,
};
type _RTConfig = typeof rtSchemaConfig;
export interface RTConfig extends _RTConfig {}

/**
 * /queue/{id} endpoint schemas
 */
export const queueSchema = z.object({
    id: z.string().min(1),
    Name: z.string().min(1),
    TicketCustomFields: typedHyperlinkSchema.array(),
    _hyperlinks: refHyperlinkSchema.array(),
});
export interface Queue extends z.infer<typeof queueSchema> {}

export const queueSchemaConfig = {
    path: `/queue/{id}`,
    getResponse: queueSchema,
};
type _QueueIdConfig = typeof queueSchemaConfig;
export interface QueueIdConfig extends _QueueIdConfig {}

/**
 * /queues endpoint schemas
 */
const queueRefSchema = typedHyperlinkSchema;
export interface QueueRef extends z.infer<typeof queueRefSchema> {}

const queuesNavCollectionSchema = createNavigatablePagedCollectionSchema(queueSchema);
// interface QueuesNavigatablePagedCollection extends z.infer<typeof queuesNavCollectionSchema> {}

export const queuesSchemaConfig = {
    path: "/queues/all",
    getResponse: queuesNavCollectionSchema,
};
type _QueuesConfig = typeof queuesSchemaConfig;
export interface QueuesConfig extends _QueuesConfig {}

/**
 * /customfields endpoint schemas
 */
const customFieldRefSchema = typedHyperlinkSchema;
export interface CustomFieldRef extends z.infer<typeof customFieldRefSchema> {}

export const customFieldsSchemaConfig = {
    path: `/customfields`,
    getResponse: createPagedCollectionSchema(customFieldRefSchema),
};
type _CustomFieldsConfig = typeof customFieldsSchemaConfig;
export interface CustomFieldsConfig extends _CustomFieldsConfig {}

/**
 * /queue/{id}/customfields endpoint schemas
 */
export const queueCustomFieldsSchemaConfig = {
    path: `/queue/{id}/customfields`,
    getResponse: createPagedCollectionSchema(customFieldRefSchema),
};
type _QueueCustomFieldsConfig = typeof queueCustomFieldsSchemaConfig;
export interface QueueCustomFieldsConfig extends _QueueCustomFieldsConfig {}

/**
 * /customfield/{id} endpoint schemas
 */
const customFieldSchema = z.object({
    id: z.string().min(1),
    Name: z.string().min(1),
    Description: z.string().min(0),
    Values: z.array(z.string().min(1)),
    Type: z.string().min(1),
    Disabled: z.enum(["0", "1"]),
    MaxValues: z.number().min(0),
    Pattern: z.string().min(0),
    EntryHint: z.string().min(0).optional(),
    BasedOn: typedHyperlinkSchema.optional(),
    Dependents: z.array(z.lazy((): ZodType => customFieldSchema)).optional(),
    _hyperlinks: refHyperlinkSchema.array(),
});
export interface CustomField extends z.infer<typeof customFieldSchema> {}

export const customFieldSchemaConfig = {
    path: `/customfield/{id}`,
    getResponse: customFieldSchema,
};
type _CustomFieldConfig = typeof customFieldSchemaConfig;
export interface CustomFieldConfig extends _CustomFieldConfig {}

/**
 * No specific endpoint, used with /queue/{id} to get queue's custom fields from 'TicketCustomFields' hyperlinks array
 */
export const queueTicketCustomFieldsSchemaConfig = {
    path: `/queue/{id}`,
    getResponse: customFieldSchema.array(),
};
type _QueueTicketCustomFieldsConfig = typeof queueTicketCustomFieldsSchemaConfig;
export interface QueueTicketCustomFieldsConfig extends _QueueTicketCustomFieldsConfig {}

/**
 * /customfield/{id}/values endpoint schemas
 */
const customFieldValueRefSchema = typedHyperlinkSchema;
export interface CustomFieldValueRef extends z.infer<typeof customFieldValueRefSchema> {}

export const customFieldValuesSchemaConfig = {
    path: `/customfield/{id}/values`,
    getResponse: createPagedCollectionSchema(customFieldValueRefSchema),
};
type _CustomFieldValuesConfig = typeof customFieldValuesSchemaConfig;
export interface CustomFieldValuesConfig extends _CustomFieldValuesConfig {}

/**
 * /customfield/{id}/value/{value_id} endpoint schemas
 */
const customFieldValueSchema = z.object({
    id: z.string().min(1),
    Name: z.string().min(1),
    Description: z.string().min(0),
    Category: z.string().min(0),
    _hyperlinks: refHyperlinkSchema.array(),
});
export interface CustomFieldValue extends z.infer<typeof customFieldValueSchema> {}

export const customFieldValueSchemaConfig = {
    path: `/customfield/{id}/value/{value_id}`,
    getResponse: customFieldValueSchema,
};
type _CustomFieldValueConfig = typeof customFieldValueSchemaConfig;
export interface CustomFieldValueConfig extends _CustomFieldValueConfig {}

/**
 * /ticket/{id} endpoint schemas
 */
const customFieldHyperlinkSchema = typedHyperlinkSchema.and(
    z.object({
        name: z.string().min(1),
        values: z.array(z.string().min(1)).optional(),
    })
);

const ticketSchema = z.object({
    id: z.string().min(1),
    Subject: z.string().min(1),
    Type: z.string().min(1),
    Status: z.string().min(1),
    Requestor: z.array(z.string().min(1)),
    InitialPriority: z.number().min(0),
    Priority: z.number().min(0),
    FinalPriority: z.number().min(0),
    TimeLeft: z.number().min(0),
    TimeWorked: z.number().min(0),
    TimeEstimated: z.number().min(0),
    Cc: z.array(z.string().min(1)),
    AdminCc: z.array(z.string().min(1)),
    Started: z
        .string()
        .min(1)
        .transform((str) => new Date(str)),
    Resolved: z
        .string()
        .min(1)
        .transform((str) => new Date(str)),
    Starts: z
        .string()
        .min(1)
        .transform((str) => new Date(str)),
    Due: z
        .string()
        .min(1)
        .transform((str) => new Date(str)),
    Created: z
        .string()
        .min(1)
        .transform((str) => new Date(str)),
    LastUpdated: z
        .string()
        .min(1)
        .transform((str) => new Date(str)),
    Queue: typedHyperlinkSchema,
    Owner: typedHyperlinkSchema,
    Creator: typedHyperlinkSchema,
    LastUpdatedBy: typedHyperlinkSchema,
    EffectiveId: typedHyperlinkSchema,
    CustomFields: z.array(customFieldHyperlinkSchema),
    _hyperlinks: z.array(refHyperlinkSchema),
});
export interface Ticket extends z.infer<typeof ticketSchema> {}

const updateTicketOptionsSchema = z
    .object({
        id: z.string().min(1),
        Subject: z.string().min(1),
        Status: z.string().min(1),
        Description: z.string().min(1), // Content Or Description?
        Content: z.string().min(1), // Content Or Description?
        Requestor: z.string().email().min(1),
        Owner: z.string().email().min(1),
        TimeWorked: z.number().min(0),
        CustomFields: z.record(z.string().min(1), z.any()),
    })
    .partial();
export interface UpdateTicketOptions extends z.infer<typeof updateTicketOptionsSchema> {}

const updateTicketSchema = z.string().array();
export interface UpdateTicket extends z.infer<typeof updateTicketSchema> {}

export const ticketIdSchemaConfig = {
    path: `/ticket/{id}`,
    getResponse: ticketSchema,
    putRequest: updateTicketOptionsSchema,
    putResponse: updateTicketSchema,
};
type _TicketIdConfig = typeof ticketIdSchemaConfig;
export interface TicketIdConfig extends _TicketIdConfig {}

/**
 * /tickets endpoint schemas
 */
const ticketRefSchema = typedHyperlinkSchema;
export interface TicketRef extends z.infer<typeof ticketRefSchema> {}

export const ticketsSchemaConfig = {
    path: `/tickets`,
    getResponse: createPagedCollectionSchema(ticketRefSchema),
};
type _TicketsConfig = typeof ticketsSchemaConfig;
export interface TicketsConfig extends _TicketsConfig {}

/**
 * /ticket endpoint schemas
 */
export const ticketSchemaConfig = {
    path: `/ticket`,
    postRequest: updateTicketOptionsSchema,
    postResponse: ticketSchema,
};
type _TicketConfig = typeof ticketSchemaConfig;
export interface TicketConfig extends _TicketConfig {}

export type CreateTicket = z.infer<typeof typedHyperlinkSchema>;

// User related types
export type UserRef = TypedHyperlinkEntity;

export interface User {
    id: string;
    Name: string;
    Email: string;
    RealName: string;
    Privileged: "0" | "1";
    _hyperlinks: RefHyperlinkEntity[];
}

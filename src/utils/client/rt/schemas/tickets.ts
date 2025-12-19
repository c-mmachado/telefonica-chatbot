import { z } from "zod";

import { HyperlinkRef, refHyperlinkSchema, TypedHyperlinkEntity, typedHyperlinkSchema } from "./base";
import {
    BaseSchemaEndpointConfigurer,
    Client,
    createSchemaEndpointConfig,
    SchemaEndpointConfig,
    SchemaEndpointConfigurer,
} from "../../base";
import {
    convertToNavigatablePagedCollection,
    createRTNavigatablePagedCollectionSchema,
    RTNavigatablePagedCollection,
} from "./pagination";

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
    Started: z.coerce.date(),
    // .string()
    // .min(1)
    // .transform((str) => new Date(str)),
    Resolved: z.coerce.date(),
    // .string()
    // .min(1)
    // .transform((str) => new Date(str)),
    Starts: z.coerce.date(),
    // .string()
    // .min(1)
    // .transform((str) => new Date(str)),
    Due: z.coerce.date(),
    // .string()
    // .min(1)
    // .transform((str) => new Date(str)),
    Created: z.coerce.date(),
    // .string()
    // .min(1)
    // .transform((str) => new Date(str)),
    LastUpdated: z.coerce.date(),
    // .string()
    // .min(1)
    // .transform((str) => new Date(str)),
    Queue: typedHyperlinkSchema,
    Owner: typedHyperlinkSchema,
    Creator: typedHyperlinkSchema,
    LastUpdatedBy: typedHyperlinkSchema,
    EffectiveId: typedHyperlinkSchema,
    CustomFields: z.array(customFieldHyperlinkSchema),
    _hyperlinks: z.array(refHyperlinkSchema),
});
export interface Ticket extends z.infer<typeof ticketSchema> {}

const updateTicketSchema = z
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
export interface UpdateTicket extends z.infer<typeof updateTicketSchema> {}

const updatedTicketSchema = z.string().array();
export interface UpdatedTicket extends z.infer<typeof updatedTicketSchema> {}

export const ticketIdSchemaConfig = createSchemaEndpointConfig({
    path: `/ticket/{id}`,
    methods: {
        get: {
            response: ticketSchema as z.ZodType<Ticket>,
        },
        put: {
            request: updateTicketSchema as z.ZodType<UpdateTicket>,
            response: updatedTicketSchema as z.ZodType<UpdatedTicket>,
        },
    },
    // getResponse: ticketSchema,
    // putRequest: updateTicketSchema,
    // putResponse: updatedTicketSchema,
});
type _TicketIdConfig = typeof ticketIdSchemaConfig;
export interface TicketIdConfig extends _TicketIdConfig {}

/**
 * /tickets endpoint schemas
 */
const ticketRefSchema = typedHyperlinkSchema;
export interface TicketRef extends z.infer<typeof ticketRefSchema> {}

export interface TicketsConfig extends SchemaEndpointConfig {
    path: `/tickets`;
    methods: {
        get: {
            response: z.ZodType<RTNavigatablePagedCollection<TicketsConfig, Ticket>>;
        };
    };
}

const ticketsNavCollectionSchema = createRTNavigatablePagedCollectionSchema<z.ZodType<Ticket>, TicketsConfig>(
    ticketSchema
);
// interface TicketsNavigatablePagedCollection extends z.infer<typeof ticketsNavCollectionSchema> {}

export const ticketsSchemaConfig = createSchemaEndpointConfig({
    path: `/tickets`,
    methods: {
        get: {
            response: ticketsNavCollectionSchema,
        },
    },
    // getResponse: createPagedCollectionSchema(ticketRefSchema),
});
// type _TicketsConfig = typeof ticketsSchemaConfig;
// export interface TicketsConfig extends _TicketsConfig {}

/**
 * /ticket endpoint schemas
 */
export const ticketSchemaConfig = createSchemaEndpointConfig({
    path: `/ticket`,
    methods: {
        post: {
            request: updateTicketSchema as z.ZodType<UpdateTicket>,
            response: typedHyperlinkSchema as z.ZodType<TypedHyperlinkEntity>,
        },
    },
});
type _TicketConfig = typeof ticketSchemaConfig;
export interface TicketConfig extends _TicketConfig {}

/**
 * /ticket/{id}/correspond endpoint schemas
 */
const correspondTicketSchema = z.object({
    Subject: z.string().min(1),
    Content: z.string().min(1),
    ContentType: z.string().min(1), //z.enum([HttpContentTypes.Html, HttpContentTypes.TextPlain]),
    TimeTaken: z.string().min(1),
    Attachments: z.array(z.any()).optional(),
});
export interface CorrespondTicket extends z.infer<typeof correspondTicketSchema> {}

const correspondedTicketSchema = z.array(z.string());
export interface CorrespondedTicket extends z.infer<typeof correspondedTicketSchema> {}

export const ticketCorrespondSchemaConfig = createSchemaEndpointConfig({
    path: `/ticket/{id}/correspond`,
    methods: {
        post: {
            request: correspondTicketSchema as z.ZodType<CorrespondTicket>,
            response: correspondedTicketSchema as z.ZodType<CorrespondedTicket>,
        },
    },
});
type _TicketCorrespondConfig = typeof ticketCorrespondSchemaConfig;
export interface TicketCorrespondConfig extends _TicketCorrespondConfig {}

/**
 *
 */
export interface TicketsEndpointConfigurer extends SchemaEndpointConfigurer<TicketsConfig> {
    id(ticketId: string): TicketIdEndpointConfigurer;

    create: TicketEndpointConfigurer;
}

export interface TicketIdEndpointConfigurer extends SchemaEndpointConfigurer<TicketIdConfig> {
    correspond: TicketCorrespondEndpointConfigurer;
}

export interface TicketEndpointConfigurer extends SchemaEndpointConfigurer<TicketConfig> {}

export interface TicketCorrespondEndpointConfigurer extends SchemaEndpointConfigurer<TicketCorrespondConfig> {}

/**
 *
 */
export class EndpointFactory {
    public static tickets(client: Client): TicketsEndpointConfigurer {
        return new DefaultTicketsEndpointConfigurer(client);
    }

    public static ticketId(client: Client, ticketId: string): TicketIdEndpointConfigurer {
        return new DefaultTicketIdEndpointConfigurer(client, ticketId);
    }

    public static ticket(client: Client): TicketEndpointConfigurer {
        return new DefaultTicketEndpointConfigurer(client);
    }

    public static ticketCorrespond(client: Client, ticketId: string): TicketCorrespondEndpointConfigurer {
        return new DefaultTicketCorrespondEndpointConfigurer(client, ticketId);
    }
}

/**
 *
 */
class DefaultTicketsEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<TicketsConfig>
    implements TicketsEndpointConfigurer
{
    constructor(client: Client) {
        super(client, ticketsSchemaConfig, {
            after: {
                get: (response: unknown) => {
                    return convertToNavigatablePagedCollection<TicketsConfig, TicketRef, Ticket>(
                        response,
                        this.client,
                        ticketsSchemaConfig,
                        this.callbacks ?? {},
                        this.path(),
                        (ref) => !!(ref?.ref === HyperlinkRef.Ticket && ref?.id),
                        async (ref) => {
                            return this.id(ref.id!).request.get();
                        }
                    );
                },
            },
        });
    }

    public id(ticketId: string): TicketIdEndpointConfigurer {
        return EndpointFactory.ticketId(this.client, ticketId);
    }

    public get create(): TicketEndpointConfigurer {
        return EndpointFactory.ticket(this.client);
    }
}

class DefaultTicketIdEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<TicketIdConfig>
    implements TicketIdEndpointConfigurer
{
    constructor(client: Client, private readonly _ticketId: string) {
        super(client, ticketIdSchemaConfig);
        this.variable("id", this._ticketId);
    }

    public get correspond(): TicketCorrespondEndpointConfigurer {
        return EndpointFactory.ticketCorrespond(this.client, this._ticketId);
    }
}

class DefaultTicketEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<TicketConfig>
    implements TicketEndpointConfigurer
{
    constructor(client: Client) {
        super(client, ticketSchemaConfig);
    }
}

class DefaultTicketCorrespondEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<TicketCorrespondConfig>
    implements TicketCorrespondEndpointConfigurer
{
    constructor(client: Client, private readonly _ticketId: string) {
        super(client, ticketCorrespondSchemaConfig);
        this.variable("id", this._ticketId);
    }
}

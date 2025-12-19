import { z } from "zod";

import { HyperlinkRef, refHyperlinkSchema, typedHyperlinkSchema } from "./base";
import {
    createSchemaEndpointConfig,
    SchemaEndpointConfig,
    BaseSchemaEndpointConfigurer,
    SchemaEndpointConfigurer,
    Client,
} from "../../base";
import {
    convertToNavigatablePagedCollection,
    createRTNavigatablePagedCollectionSchema,
    RTNavigatablePagedCollection,
} from "./pagination";
import {
    customFieldSchema,
    CustomField,
    CustomFieldIdEndpointConfigurer,
    EndpointFactory as CFEndpointFactory,
} from "./customFields";

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

export const queueSchemaConfig = createSchemaEndpointConfig({
    path: `/queue/{id}`,
    methods: {
        get: {
            response: queueSchema as z.ZodType<Queue>,
        },
    },
});
type _QueueIdConfig = typeof queueSchemaConfig;
export interface QueueIdConfig extends _QueueIdConfig {}

/**
 * /queues endpoint schemas
 */
const queueRefSchema = typedHyperlinkSchema;
export interface QueueRef extends z.infer<typeof queueRefSchema> {}

export interface QueuesConfig extends SchemaEndpointConfig {
    path: "/queues/all";
    methods: {
        get: {
            response: z.ZodType<RTNavigatablePagedCollection<QueuesConfig, Queue>>;
        };
    };
}

const queuesNavCollectionSchema = createRTNavigatablePagedCollectionSchema<z.ZodType<Queue>, QueuesConfig>(queueSchema);
// interface QueuesNavigatablePagedCollection extends z.infer<typeof queuesNavCollectionSchema> {}

export const queuesSchemaConfig: QueuesConfig = createSchemaEndpointConfig({
    path: "/queues/all",
    methods: {
        get: {
            response: queuesNavCollectionSchema,
        },
    },
});
// type _QueuesConfig = typeof queuesSchemaConfig;
// export interface QueuesConfig extends _QueuesConfig {}

/**
 * /queue/{id}/customfields endpoint schemas
 */
// export interface QueueCustomFieldsConfig extends SchemaEndpointConfig {
//     path: `/queue/{id}/customfields`;
//     methods: {
//         get: {
//             response: z.ZodType<RTNavigatablePagedCollection<QueueCustomFieldsConfig, CustomField>>;
//         };
//     };
// }

// const queueCustomFieldsNavCollectionSchema = createNavigatablePagedCollectionSchema<
//     z.ZodType<CustomField>,
//     QueueCustomFieldsConfig
// >(customFieldSchema);

// export const queueCustomFieldsSchemaConfig: QueueCustomFieldsConfig = createSchemaEndpointConfig({
//     path: `/queue/{id}/customfields`,
//     methods: {
//         get: {
//             response: queueCustomFieldsNavCollectionSchema,
//         },
//     },
// });
// type _QueueCustomFieldsConfig = typeof queueCustomFieldsSchemaConfig;
// export interface QueueCustomFieldsConfig extends _QueueCustomFieldsConfig {}

/**
 * No specific endpoint, used with /queue/{id} to get queue's custom fields from 'TicketCustomFields' hyperlinks array
 */
export const queueTicketCustomFieldsSchemaConfig = createSchemaEndpointConfig({
    path: `/queue/{id}`,
    methods: {
        get: {
            response: (customFieldSchema as z.ZodType<CustomField>).array(),
        },
    },
});
type _QueueTicketCustomFieldsConfig = typeof queueTicketCustomFieldsSchemaConfig;
export interface QueueTicketCustomFieldsConfig extends _QueueTicketCustomFieldsConfig {}

/**
 *
 */
export interface QueuesEndpointConfigurer extends SchemaEndpointConfigurer<QueuesConfig> {
    id(queueId: string): QueueIdEndpointConfigurer;
}

export interface QueueIdEndpointConfigurer extends SchemaEndpointConfigurer<QueueIdConfig> {
    ticketCustomFields: QueueTicketCustomFieldsEndpointConfigurer;
}

export interface QueueTicketCustomFieldsEndpointConfigurer
    extends SchemaEndpointConfigurer<QueueTicketCustomFieldsConfig> {
    id(customFieldId: string): CustomFieldIdEndpointConfigurer;
}

// export interface QueueCustomFieldsEndpointConfigurer extends SchemaEndpointConfigurer<QueueCustomFieldsConfig> {
//     id(customFieldId: string): CustomFieldIdEndpointConfigurer;
// }

/**
 *
 */
export class EndpointFactory {
    public static queues(client: Client): QueuesEndpointConfigurer {
        return DefaultQueuesEndpointConfigurer.create(client);
    }

    public static queueId(client: Client, queueId: string): QueueIdEndpointConfigurer {
        return DefaultQueueIdEndpointConfigurer.create(client, queueId);
    }

    // public static queueCustomFields(client: Client): QueueCustomFieldsEndpointConfigurer {
    //     return DefaultQueueCustomFieldsEndpointConfigurer.create(client);
    // }

    public static queueTicketCustomFields(client: Client, queueId: string): QueueTicketCustomFieldsEndpointConfigurer {
        return DefaultQueueTicketCustomFieldsEndpointConfigurer.create(client, queueId);
    }
}

/**
 *
 */
class DefaultQueuesEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<QueuesConfig>
    implements QueuesEndpointConfigurer
{
    public static create(client: Client): QueuesEndpointConfigurer {
        return new DefaultQueuesEndpointConfigurer(client);
    }

    private constructor(client: Client) {
        super(client, queuesSchemaConfig, {
            after: {
                get: async (response: unknown) => {
                    return convertToNavigatablePagedCollection<QueuesConfig, QueueRef, Queue>(
                        response,
                        this.client,
                        queuesSchemaConfig,
                        this.callbacks ?? {},
                        this.path(),
                        (ref) => !!(ref?.ref === HyperlinkRef.Queue && ref?.id),
                        async (ref) => {
                            return this.id(ref.id!).request.get();
                        }
                    );
                },
                // if (!response || typeof response !== "object") {
                //     return Promise.reject(new Error(`Invalid response format for resource '${this.path()}'`));
                // }
                // const page = response as RTPagedCollection<QueueRef>;
                // const refs = page.items ?? [];
                // const queues: Queue[] = await Promise.all(
                //     refs
                //         .filter((ref) => ref.ref === HyperlinkRef.Queue && ref.id)
                //         // .map((ref) => rt.queues.id(ref.id!).request.get())
                //         .map((ref) => this.id(ref.id!).request.get())
                // );
                // const newPage: RTPagedCollection<Queue> = {
                //     ...page,
                //     items: queues,
                // };
                // return Promise.resolve(
                //     createRTNavigatablePagedCollection<QueuesConfig, Queue>(
                //         this.client,
                //         queuesSchemaConfig,
                //         newPage
                //     )
                // );
            },
        });
    }

    public id(queueId: string): QueueIdEndpointConfigurer {
        return EndpointFactory.queueId(this.client, queueId);
    }
}

class DefaultQueueIdEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<QueueIdConfig>
    implements QueueIdEndpointConfigurer
{
    public static create(client: Client, queueId: string): QueueIdEndpointConfigurer {
        return new DefaultQueueIdEndpointConfigurer(client, queueId);
    }

    private constructor(client: Client, private readonly _queueId: string) {
        super(client, queueSchemaConfig);
        this.variable("id", this._queueId);
    }

    public get ticketCustomFields(): QueueTicketCustomFieldsEndpointConfigurer {
        return EndpointFactory.queueTicketCustomFields(this.client, this._queueId);
    }
}

// class DefaultQueueCustomFieldsEndpointConfigurer
//     extends BaseSchemaEndpointConfigurer<QueueCustomFieldsConfig>
//     implements QueueCustomFieldsEndpointConfigurer
// {
//     public static create(client: Client): QueueCustomFieldsEndpointConfigurer {
//         return new DefaultQueueCustomFieldsEndpointConfigurer(client);
//     }

//     private constructor(client: Client) {
//         super(client, queueCustomFieldsSchemaConfig, {
//             after: {
//                 get: async (response: unknown) => {
//                     return convertToNavigatablePagedCollection<QueueCustomFieldsConfig, HyperlinkEntity, CustomField>(
//                         response,
//                         this.client,
//                         queueCustomFieldsSchemaConfig,
//                         this.path(),
//                         (ref) => !!(ref?.ref === HyperlinkRef.CustomField && ref?.id),
//                         async (ref) => {
//                             return this.id(ref.id!).request.get();
//                         }
//                     );
//                 },
//             },
//         });
//     }

//     public id(customFieldId: string): CustomFieldIdEndpointConfigurer {
//         return CFEndpointFactory.customFieldId(this.client, customFieldId);
//     }
// }

class DefaultQueueTicketCustomFieldsEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<QueueTicketCustomFieldsConfig>
    implements QueueTicketCustomFieldsEndpointConfigurer
{
    public static create(client: Client, queueId: string): QueueTicketCustomFieldsEndpointConfigurer {
        return new DefaultQueueTicketCustomFieldsEndpointConfigurer(client, queueId);
    }

    private constructor(client: Client, private readonly _queueId: string) {
        super(client, queueTicketCustomFieldsSchemaConfig, {
            after: {
                get: async (response: unknown): Promise<CustomField[]> => {
                    if (!response || typeof response !== "object") {
                        return Promise.reject(new Error(`Invalid response format for resource '${this.path()}'`));
                    }
                    const refs = (response as Queue).TicketCustomFields ?? [];
                    const fields = [];
                    for (const ref of refs) {
                        if (ref.ref === HyperlinkRef.CustomField && ref.id) {
                            // fields.push(await rt.customFields.id(ref.id).request.get());
                            fields.push(await this.id(ref.id).request.get());
                        }
                    }
                    return Promise.resolve(fields);
                },
            },
        });
        this.variable("id", this._queueId);
    }

    public id(customFieldId: string): CustomFieldIdEndpointConfigurer {
        return CFEndpointFactory.customFieldId(this.client, customFieldId);
    }
}

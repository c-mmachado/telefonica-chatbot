import { z } from "zod";

import {
    BaseSchemaEndpointConfigurer,
    Client,
    createSchemaEndpointConfig,
    SchemaEndpointConfig,
    SchemaEndpointConfigurer,
} from "../../base";
import { HyperlinkRef, refHyperlinkSchema, typedHyperlinkSchema } from "./base";
import {
    convertToNavigatablePagedCollection,
    createRTNavigatablePagedCollectionSchema,
    RTNavigatablePagedCollection,
} from "./pagination";

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

export const customFieldValueSchemaConfig = createSchemaEndpointConfig({
    path: `/customfield/{id}/value/{value_id}`,
    methods: {
        get: {
            response: customFieldValueSchema as z.ZodType<CustomFieldValue>,
        },
    },
});
type _CustomFieldValueConfig = typeof customFieldValueSchemaConfig;
export interface CustomFieldValueConfig extends _CustomFieldValueConfig {}

/**
 * /customfield/{id}/values endpoint schemas
 */
const customFieldValueRefSchema = typedHyperlinkSchema;
export interface CustomFieldValueRef extends z.infer<typeof customFieldValueRefSchema> {}

export interface CustomFieldValuesConfig extends SchemaEndpointConfig {
    path: `/customfield/{id}/values`;
    methods: {
        get: {
            response: z.ZodType<RTNavigatablePagedCollection<CustomFieldValuesConfig, CustomFieldValue>>;
        };
    };
}

const customFieldValuesNavCollectionSchema = createRTNavigatablePagedCollectionSchema<
    z.ZodType<CustomFieldValue>,
    CustomFieldValuesConfig
>(customFieldValueSchema);
// interface CustomFieldValuesNavigatablePagedCollection extends z.infer<typeof customFieldValuesNavCollectionSchema> {}

export const customFieldValuesSchemaConfig = createSchemaEndpointConfig({
    path: `/customfield/{id}/values`,
    methods: {
        get: {
            response: customFieldValuesNavCollectionSchema,
        },
    },
});
// type _CustomFieldValuesConfig = typeof customFieldValuesSchemaConfig;
// export interface CustomFieldValuesConfig extends _CustomFieldValuesConfig {}

/**
 *
 */
export interface CustomFieldValuesEndpointConfigurer extends SchemaEndpointConfigurer<CustomFieldValuesConfig> {
    id(customFieldValueId: string): CustomFieldValueIdEndpointConfigurer;
}

export interface CustomFieldValueIdEndpointConfigurer extends SchemaEndpointConfigurer<CustomFieldValueConfig> {}

/**
 *
 */
export class EndpointFactory {
    public static customFieldValues(client: Client, customFieldId: string): CustomFieldValuesEndpointConfigurer {
        return DefaultCustomFieldValuesEndpointConfigurer.create(client, customFieldId);
    }

    public static customFieldValueId(
        client: Client,
        customFieldId: string,
        customFieldValueId: string
    ): CustomFieldValueIdEndpointConfigurer {
        return DefaultCustomFieldValueIdEndpointConfigurer.create(client, customFieldId, customFieldValueId);
    }
}

/**
 *
 */
class DefaultCustomFieldValuesEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<CustomFieldValuesConfig>
    implements CustomFieldValuesEndpointConfigurer
{
    public static create(client: Client, customFieldId: string): CustomFieldValuesEndpointConfigurer {
        return new DefaultCustomFieldValuesEndpointConfigurer(client, customFieldId);
    }

    private constructor(client: Client, private readonly _customFieldId: string) {
        super(client, customFieldValuesSchemaConfig, {
            after: {
                get: (response: unknown) => {
                    return convertToNavigatablePagedCollection<
                        CustomFieldValuesConfig,
                        CustomFieldValueRef,
                        CustomFieldValue
                    >(
                        response,
                        this.client,
                        customFieldValuesSchemaConfig,
                        this.callbacks ?? {},
                        this.path(),
                        (ref) => !!(ref?.ref === HyperlinkRef.CustomFieldValue && ref?.id),
                        async (ref) => {
                            return this.id(ref.id!).request.get();
                        }
                    );
                },
            },
        });
        this.variable("id", this._customFieldId);
    }

    public id(customFieldValueId: string): CustomFieldValueIdEndpointConfigurer {
        return EndpointFactory.customFieldValueId(this.client, this._customFieldId, customFieldValueId);
    }
}

class DefaultCustomFieldValueIdEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<CustomFieldValueConfig>
    implements CustomFieldValueIdEndpointConfigurer
{
    public static create(
        client: Client,
        customFieldId: string,
        customFieldValueId: string
    ): CustomFieldValueIdEndpointConfigurer {
        return new DefaultCustomFieldValueIdEndpointConfigurer(client, customFieldId, customFieldValueId);
    }

    private constructor(
        client: Client,
        private readonly _customFieldId: string,
        private readonly _customFieldValueId: string
    ) {
        super(client, customFieldValueSchemaConfig);
        this.variable("id", this._customFieldId);
        this.variable("value_id", this._customFieldValueId);
    }
}

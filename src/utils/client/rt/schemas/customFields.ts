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
import { CustomFieldValuesEndpointConfigurer, EndpointFactory as CFVEndpointFactory } from "./customFieldValues";

/**
 * /customfield/{id} endpoint schemas
 */
export const customFieldSchema = z.object({
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
    Dependents: z.array(z.lazy((): z.ZodType => customFieldSchema)).optional(),
    _hyperlinks: refHyperlinkSchema.array(),
});
export interface CustomField extends z.infer<typeof customFieldSchema> {}

export const customFieldSchemaConfig = createSchemaEndpointConfig({
    path: `/customfield/{id}`,
    methods: {
        get: {
            response: customFieldSchema as z.ZodType<CustomField>,
        },
    },
    // getResponse: customFieldSchema,
});
type _CustomFieldConfig = typeof customFieldSchemaConfig;
export interface CustomFieldConfig extends _CustomFieldConfig {}

/**
 * /customfields endpoint schemas
 */
const customFieldRefSchema = typedHyperlinkSchema;
export interface CustomFieldRef extends z.infer<typeof customFieldRefSchema> {}

export interface CustomFieldsConfig extends SchemaEndpointConfig {
    path: `/customfields`;
    methods: {
        get: {
            response: z.ZodType<RTNavigatablePagedCollection<CustomFieldsConfig, CustomField>>;
        };
    };
}

const customFieldsNavCollectionSchema = createRTNavigatablePagedCollectionSchema<
    z.ZodType<CustomField>,
    CustomFieldsConfig
>(customFieldSchema);
// interface CustomFieldsNavigatablePagedCollection extends z.infer<typeof customFieldsNavCollectionSchema> {}

export const customFieldsSchemaConfig = createSchemaEndpointConfig({
    path: `/customfields`,
    methods: {
        get: {
            response: customFieldsNavCollectionSchema,
        },
    },
});
// type _CustomFieldsConfig = typeof customFieldsSchemaConfig;
// export interface CustomFieldsConfig extends _CustomFieldsConfig {}

/**
 *
 */
export interface CustomFieldsEndpointConfigurer extends SchemaEndpointConfigurer<CustomFieldsConfig> {
    id(customFieldId: string): CustomFieldIdEndpointConfigurer;
}

export interface CustomFieldIdEndpointConfigurer extends SchemaEndpointConfigurer<CustomFieldConfig> {
    customFieldValues: CustomFieldValuesEndpointConfigurer;
}

/**
 *
 */
export class EndpointFactory {
    public static customFields(client: Client): CustomFieldsEndpointConfigurer {
        return DefaultCustomFieldsEndpointConfigurer.create(client);
    }

    public static customFieldId(client: Client, customFieldId: string): CustomFieldIdEndpointConfigurer {
        return DefaultCustomFieldIdEndpointConfigurer.create(client, customFieldId);
    }
}

/**
 *
 */
class DefaultCustomFieldsEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<CustomFieldsConfig>
    implements CustomFieldsEndpointConfigurer
{
    public static create(client: Client): CustomFieldsEndpointConfigurer {
        return new DefaultCustomFieldsEndpointConfigurer(client);
    }

    private constructor(client: Client) {
        super(client, customFieldsSchemaConfig, {
            after: {
                get: (response: unknown) => {
                    return convertToNavigatablePagedCollection<CustomFieldsConfig, CustomFieldRef, CustomField>(
                        response,
                        this.client,
                        customFieldsSchemaConfig,
                        this.callbacks ?? {},
                        this.path(),
                        (ref) => !!(ref?.ref === HyperlinkRef.CustomField && ref?.id),
                        async (ref) => {
                            return this.id(ref.id!).request.get();
                        }
                    );
                },
            },
        });
    }

    public id(customFieldId: string): CustomFieldIdEndpointConfigurer {
        return EndpointFactory.customFieldId(this.client, customFieldId);
    }
}

class DefaultCustomFieldIdEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<CustomFieldConfig>
    implements CustomFieldIdEndpointConfigurer
{
    public static create(client: Client, customFieldId: string): CustomFieldIdEndpointConfigurer {
        return new DefaultCustomFieldIdEndpointConfigurer(client, customFieldId);
    }

    private constructor(client: Client, private _customFieldId: string) {
        super(client, customFieldSchemaConfig);
        this.variable("id", this._customFieldId);
    }

    public get customFieldValues(): CustomFieldValuesEndpointConfigurer {
        return CFVEndpointFactory.customFieldValues(this.client, this._customFieldId);
    }
}

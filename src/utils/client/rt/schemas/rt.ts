import { z } from "zod";

import { BaseSchemaEndpointConfigurer, Client, createSchemaEndpointConfig, SchemaEndpointConfigurer } from "../../base";

/**
 * /rt endpoint schemas
 */
const rtGetResponseSchema = z.object({
    Version: z.string().min(1),
});
interface RT extends z.infer<typeof rtGetResponseSchema> {}
export const rtSchemaConfig = createSchemaEndpointConfig({
    path: "/rt",
    methods: {
        get: {
            response: rtGetResponseSchema as z.ZodType<RT>,
        },
    },
});
type _RTConfig = typeof rtSchemaConfig;
export interface RTConfig extends _RTConfig {}

/**
 *
 */
export interface RootEndpointConfigurer extends SchemaEndpointConfigurer<RTConfig> {}

/**
 *
 */
export class EndpointFactory {
    public static rt(client: Client): RootEndpointConfigurer {
        return new DefaultRootEndpointConfigurer(client);
    }
}

/**
 *
 */
class DefaultRootEndpointConfigurer extends BaseSchemaEndpointConfigurer<RTConfig> implements RootEndpointConfigurer {
    constructor(client: Client) {
        super(client, rtSchemaConfig);
    }
}

import { z } from "zod";

import { HyperlinkRef, refHyperlinkSchema, typedHyperlinkSchema } from "./base";
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
 * /user/{id} endpoint schemas
 */
const userSchema = z.object({
    id: z.string().min(1),
    Name: z.string().min(1),
    Email: z.string().min(1).email(),
    RealName: z.string().min(1),
    Privileged: z.enum(["0", "1"]),
    _hyperlinks: refHyperlinkSchema.array(),
});
export interface User extends z.infer<typeof userSchema> {}

const userIdSchemaConfig = createSchemaEndpointConfig({
    path: `/user/{id}`,
    methods: {
        get: {
            response: userSchema as z.ZodType<User>,
        },
    },
});
type _UserIdConfig = typeof userIdSchemaConfig;
interface UserIdConfig extends _UserIdConfig {}

/**
 * /users endpoint schemas
 */
const userRefSchema = typedHyperlinkSchema;
export interface UserRef extends z.infer<typeof userRefSchema> {}

interface UsersConfig extends SchemaEndpointConfig {
    path: `/users`;
    methods: {
        get: {
            response: z.ZodType<RTNavigatablePagedCollection<UsersConfig, User>>;
        };
    };
}

const usersNavCollectionSchema = createRTNavigatablePagedCollectionSchema<z.ZodType<User>, UsersConfig>(userSchema);
// interface UsersNavigatablePagedCollection extends z.infer<typeof usersNavCollectionSchema> {}

const usersSchemaConfig: UsersConfig = createSchemaEndpointConfig({
    path: `/users`,
    methods: {
        get: {
            response: usersNavCollectionSchema,
        },
    },
});

/**
 *
 */
export interface UsersEndpointConfigurer extends SchemaEndpointConfigurer<UsersConfig> {
    id(userId: string): UserIdEndpointConfigurer;
}

export interface UserIdEndpointConfigurer extends SchemaEndpointConfigurer<UserIdConfig> {}

/**
 *
 */
export class EndpointFactory {
    public static users(client: Client): UsersEndpointConfigurer {
        return DefaultUsersEndpointConfigurer.create(client);
    }

    public static userId(client: Client, userId: string): UserIdEndpointConfigurer {
        return DefaultUserIdEndpointConfigurer.create(client, userId);
    }
}

/**
 *
 */
class DefaultUsersEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<UsersConfig>
    implements UsersEndpointConfigurer
{
    public static create(client: Client): UsersEndpointConfigurer {
        return new DefaultUsersEndpointConfigurer(client);
    }

    private constructor(client: Client) {
        super(client, usersSchemaConfig, {
            after: {
                get: async (response: unknown) => {
                    return convertToNavigatablePagedCollection<UsersConfig, UserRef, User>(
                        response,
                        this.client,
                        this.config,
                        this.callbacks ?? {},
                        this.path(),
                        (ref) => !!(ref?.ref === HyperlinkRef.User && ref?.id),
                        async (ref) => {
                            return this.id(ref.id!).request.get();
                        }
                    );
                },
            },
        });
    }

    public id(userId: string): UserIdEndpointConfigurer {
        return EndpointFactory.userId(this.client, userId);
    }
}

class DefaultUserIdEndpointConfigurer
    extends BaseSchemaEndpointConfigurer<UserIdConfig>
    implements UserIdEndpointConfigurer
{
    public static create(client: Client, userId: string): UserIdEndpointConfigurer {
        return new DefaultUserIdEndpointConfigurer(client, userId);
    }

    private constructor(client: Client, userId: string) {
        super(client, userIdSchemaConfig);
        this.variable("id", userId);
    }
}

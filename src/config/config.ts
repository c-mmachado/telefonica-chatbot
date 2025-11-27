// import { readFileSync } from "fs";

import { z } from "zod";
import { config as dotEnv } from "dotenv";

import { envArray, envBoolean, envInt } from "../utils/schemas";
// import { recursiveMask } from "../utils/misc";

console.debug(`Loading configuration...`);
console.debug(`currentWorkingDir:`, process.cwd());

// Load environment variables from .env file.
// This only applies when running the application through 'npm run start' in a local environment as
// the Teams Toolkit will automatically load environment when running the application through it
dotEnv({
    path: "./env/.env.local",
    debug: true,
    encoding: "utf8",
    override: true, // Override existing environment variables to allow for local hotswapping
});

// const sslCertSchema = z.object({
//   key: z.instanceof(Buffer),
//   cert: z.instanceof(Buffer),
// });

const jwtOptionsSchema = z.object({
    secret: z.string().min(8).default("default_jwt_secret_change_me"),
    rootUsername: z.string().min(1).default("root"),
    rootPassword: z.string().min(1).default("password"),
});

const serverSchema = z.object({
    port: envInt({ min: 1, max: 65535, default: 3978 }),
    jwt: jwtOptionsSchema,
});

const botTypeSchema = z.enum(["MultiTenant", "SingleTenant"]);
export type BotType = z.infer<typeof botTypeSchema>;

const configSchema = z.object({
    // Azure bot settings
    botId: z.string().min(1),
    botPassword: z.string().min(1),
    botDomain: z.string().min(1),
    botType: botTypeSchema,
    botConnectionName: z.string().min(1),

    // AAD app settings
    clientId: z.string().min(1),
    tenantId: z.string().min(1),
    clientSecret: z.string().min(1),
    authority: z.string().min(1),
    authorityHost: z.string().min(1),
    scopes: envArray<z.ZodString>(z.string()).optional(),

    // Teams app settings
    teamsAppId: z.string().min(1),
    teamsAppCatalogId: z.string().optional(),
    teamsAppTenantId: z.string().min(1),

    // RT API settings
    apiEndpoint: z.string().min(1),
    apiUsername: z.string().min(1),
    apiPassword: z.string().min(1),
    apiBasePath: z.string().min(1).default("/REST/2.0"),

    // Database settings
    dbHost: z.string().min(1).default("localhost"),
    dbPort: envInt({ min: 1, max: 65535, default: 5432 }),
    dbUser: z.string().min(1).default("postgres"),
    dbPassword: z.string().min(1).default("postgres"),
    dbName: z.string().min(1).default("ticket_bot"),

    // Graph settings
    graphUsername: z.string().min(1),
    graphPassword: z.string().min(1),

    // Bot settings
    allowAll: envBoolean(false),
    botTimeout: envInt({ min: 1, default: 60 * 60 }), // in seconds

    // SSL settings
    // ssl: sslCertSchema,

    // Server settings
    server: serverSchema,
});

export type BotConfiguration = z.infer<typeof configSchema>;

function mapEnvToObj(env: NodeJS.ProcessEnv): any {
    return {
        // Azure bot settings
        botId: env.BOT_ID,
        botPassword: env.BOT_PASSWORD,
        botDomain: env.BOT_DOMAIN,
        botType: env.BOT_TYPE,
        botConnectionName: env.BOT_CONNECTION_NAME,

        // AAD app settings
        clientId: env.AAD_APP_CLIENT_ID,
        tenantId: env.AAD_APP_TENANT_ID,
        clientSecret: env.AAD_APP_CLIENT_SECRET,
        authority: env.AAD_APP_OAUTH_AUTHORITY,
        authorityHost: env.AAD_APP_OAUTH_AUTHORITY_HOST,
        scopes: env.AAD_APP_SCOPES,

        // Teams app settings
        teamsAppId: env.TEAMS_APP_ID,
        teamsAppCatalogId: env.TEAMS_APP_CATALOG_ID,
        teamsAppTenantId: env.TEAMS_APP_TENANT_ID,

        // API settings
        apiEndpoint: env.API_ENDPOINT,
        apiUsername: env.API_USERNAME,
        apiPassword: env.API_PASSWORD,
        apiBasePath: env.API_BASE_PATH,

        // Database settings
        dbHost: env.DB_HOST,
        dbPort: env.DB_PORT,
        dbUser: env.DB_USER,
        dbPassword: env.DB_PASSWORD,
        dbName: env.DB_NAME,

        // Graph settings
        graphUsername: env.GRAPH_USERNAME,
        graphPassword: env.GRAPH_PASSWORD,

        // Bot settings
        allowAll: env.ALLOW_ALL,
        botTimeout: env.BOT_TIMEOUT,

        // Server settings
        server: {
            port: env.PORT,
            // JWT settings
            jwt: {
                secret: env.JWT_SECRET,
                rootUsername: env.JWT_ROOT_USERNAME,
                rootPassword: env.JWT_ROOT_PASSWORD,
            },
        },
    };
}

export const config: BotConfiguration = configSchema.parse(mapEnvToObj(process.env));

// Create a safe version of the config object to avoid logging sensitive information
// like passwords or secrets
console.debug(`config:`, config);

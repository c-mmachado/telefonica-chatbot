import { config as dotEnv } from "dotenv";

console.debug(`[config][DEBUG] Loading configuration...`);
console.debug(`[config][DEBUG] currentWorkingDir: ${process.cwd()}`);

// Load environment variables from .env file.
// This only applies when running the application through 'npm run start' in a local environment as
// the Teams Toolkit will automatically load environment when running the application through it
dotEnv({
  path: "./env/.env.local",
  debug: true,
  encoding: "utf8",
  override: true,
});

export declare interface BotConfiguration {
  botId: string;
  botPassword: string;
  botDomain: string;
  botType: string;
  botConnectionName: string;

  clientId: string;
  tenantId: string;
  clientSecret: string;
  authorityHost: string;

  teamsAppId: string;
  teamsAppCatalogId: string;
  teamsAppTenantId: string;

  apiEndpoint: string;
  apiUsername: string;
  apiPassword: string;

  dbHost: string;
  dbPort: number;
  dbUser: string;
  dbPassword: string;
  dbName: string;
}

export const config: BotConfiguration = {
  // Azure bot settings
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  botDomain: process.env.BOT_DOMAIN,
  botType: process.env.BOT_TYPE,
  botConnectionName: process.env.BOT_CONNECTION_NAME,

  // AAD app settings
  clientId: process.env.AAD_APP_CLIENT_ID,
  tenantId: process.env.AAD_APP_TENANT_ID,
  clientSecret: process.env.AAD_APP_CLIENT_SECRET,
  authorityHost: process.env.AAD_APP_OAUTH_AUTHORITY_HOST,

  // Teams app settings
  teamsAppId: process.env.TEAMS_APP_ID,
  teamsAppCatalogId: process.env.TEAMS_APP_CATALOG_ID,
  teamsAppTenantId: process.env.TEAMS_APP_TENANT_ID,

  // API settings
  apiEndpoint: process.env.API_ENDPOINT,
  apiUsername: process.env.API_USERNAME,
  apiPassword: process.env.API_PASSWORD,

  // Database settings
  dbHost: process.env.DB_HOST,
  dbPort: parseInt(process.env.DB_PORT),
  dbUser: process.env.DB_USER,
  dbPassword: process.env.DB_PASSWORD,
  dbName: process.env.DB_NAME,
};

console.debug(`[config][DEBUG] config:\n${JSON.stringify(config, null, 2)}`);

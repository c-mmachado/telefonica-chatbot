import { config as dotEnv } from "dotenv";

console.debug(`[config][DEBUG] Loading configuration...`);
console.debug(`[config][DEBUG] currentWorkingDir: ${process.cwd()}`);

// Load environment variables from .env file.
// This only applies when running the application through 'npm run start' in a local environment as
// the Teams Toolkit will automatically load environment when running the application through 
dotEnv({
  path: "./env/.env.local",
  debug: true,
  encoding: "utf8",
  override: true,
});

const config = {
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
};

console.debug(`[config][DEBUG] config:\n${JSON.stringify(config, null, 2)}`);

export default config;

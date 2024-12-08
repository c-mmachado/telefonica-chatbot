import config from "../config/config";

export interface JsonObject {
  [key: string]: any;
}

export interface TokenExchangeResponse {
  token_type: string;
  expires_in: number;
  ext_expires_in: number;
  access_token: string;
}

export class AppInstallUtils {
  private static getAccessToken(
    tenantId: string
  ): Promise<TokenExchangeResponse> {
    return new Promise(
      async (
        resolve: (value: TokenExchangeResponse) => void,
        reject: (reason: Error) => void
      ) => {
        fetch(
          `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
          {
            method: "POST",
            headers: {
              "Content-Type": "application/x-www-form-urlencoded",
            },
            body: new URLSearchParams(
              Object.entries({
                grant_type: "client_credentials",
                scope: "https://graph.microsoft.com/.default",
                client_id: config.clientId,
                client_secret: config.clientSecret,
              })
            ).toString(),
          }
        )
          .then((response: Response): Promise<TokenExchangeResponse> => {
            return response?.json();
          })
          .then((response: TokenExchangeResponse): void => {
            console.debug(
              `[${
                AppInstallUtils.name
              }] getAccessToken response:\n${JSON.stringify(
                response,
                null,
                2
              )}\n`
            );
            resolve(response);
          })
          .catch((error: Error): void => {
            console.debug(
              `${AppInstallUtils.name} getAccessToken error:\n${JSON.stringify(
                error,
                null,
                2
              )}\n`
            );
            reject(error);
          });
      }
    );
  }

  public static async installAppInPersonalScope(
    tenantId: string,
    userId: string
  ): Promise<Response> {
    return new Promise(
      async (
        resolve: (value: Response) => void,
        reject: (reason: Error) => void
      ) => {
        console.debug(
          `[${AppInstallUtils.name}] installAppInPersonalScope tenantId: ${tenantId}`
        );
        console.debug(
          `[${AppInstallUtils.name}] installAppInPersonalScope userId: ${userId}`
        );

        const token = await AppInstallUtils.getAccessToken(tenantId);
        fetch(
          `https://graph.microsoft.com/v1.0/users/${userId}/teamwork/installedApps`,
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: `${token.token_type} ${token.access_token}`,
            },
            // body: JSON.stringify({
            //   "teamsApp@odata.bind":
            //     "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" +
            //     config.teamsAppCatalogId,
            // }),
          }
        )
          .then((response: Response): void => {
            console.debug(
              `[${
                AppInstallUtils.name
              }] installAppInPersonalScope response:\n${JSON.stringify(
                response,
                null,
                2
              )}\n`
            );
            resolve(response);
          })
          .catch(async (error: Error): Promise<void> => {
            console.debug(
              `[${
                AppInstallUtils.name
              }] installAppInPersonalScope error:\n${JSON.stringify(
                error,
                null,
                2
              )}\n`
            );
            await AppInstallUtils.triggerConversationUpdate(tenantId, userId);
            reject(error);
          });
      }
    );
  }

  private static async triggerConversationUpdate(
    tenantId: string,
    userId: string
  ) {
    return new Promise(async (resolve: (value: number | Error) => void) => {
      const token = await AppInstallUtils.getAccessToken(tenantId);

      fetch(
        `https://graph.microsoft.com/v1.0/users/${userId}/teamwork/installedApps?$expand=teamsApp,teamsAppDefinition&$filter=teamsApp/externalId eq '${config.teamsAppId}'`,
        {
          method: "GET",
          headers: {
            Authorization: `${token.token_type} ${token.access_token}`,
          },
        }
      )
        .then(async (response: Response): Promise<JsonObject> => {
          console.debug(
            `[${
              AppInstallUtils.name
            }] triggerConversationUpdate response:\n${JSON.stringify(
              response,
              null,
              2
            )}\n`
          );
          return response?.json();
        })
        .then(async (response: JsonObject): Promise<void> => {
          const installedAppsMap = response?.value?.map(
            (element: JsonObject) => element.teamsApp.externalId
          );
          if (!!installedAppsMap && installedAppsMap.length > 0) {
            installedAppsMap.value.forEach(async (apps: JsonObject) => {
              let result = await AppInstallUtils.installAppInPersonalChatScope(
                `${token.token_type} ${token.access_token}`,
                userId,
                apps.id
              );
            });
          }
          resolve(response.status);
        })
        .catch((error: Error): void => {
          console.debug(
            `[${
              AppInstallUtils.name
            }] triggerConversationUpdate error:\n${JSON.stringify(
              error,
              null,
              2
            )}\n`
          );
          resolve(error);
        });
    });
  }

  private static async installAppInPersonalChatScope(
    accessToken: string,
    userId: string,
    appId: string
  ) {
    return new Promise(async (resolve: (value: number | Error) => void) => {
      fetch(
        `https://graph.microsoft.com/v1.0/users/${userId}/teamwork/installedApps/${appId}/chat`,
        {
          method: "GET",
          headers: {
            Authorization: accessToken,
          },
        }
      )
        .then((response: Response): void => {
          console.debug(
            `[${
              AppInstallUtils.name
            }] installAppInPersonalChatScope response:\n${JSON.stringify(
              response,
              null,
              2
            )}\n`
          );
          resolve(response.status);
        })
        .catch((error: Error): void => {
          resolve(error);
        });
    });
  }
}

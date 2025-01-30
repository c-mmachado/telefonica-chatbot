@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@maxLength(42)
param botDisplayName string

param botServiceName string = resourceBaseName
param botServiceSku string = 'F0'
param botDomain string
param botType string
param botTenantId string
param botLoginUrl string
param botAadAppClientId string

param aadAppClientId string
@secure()
param aadAppClientSecret string

// Register your web service as a bot with the Bot Framework
resource botService 'Microsoft.BotService/botServices@2021-03-01' = {
  kind: 'azurebot'
  location: 'global'
  name: botServiceName
  properties: {
    displayName: botDisplayName
    endpoint: 'https://${botDomain}/api/messages'
    msaAppId: botAadAppClientId
    msaAppType: botType
    msaAppTenantId: botType == 'SingleTenant' ? botTenantId : ''
  }
  sku: {
    name: botServiceSku
  }
}

// Connect the bot service to Microsoft Teams
resource botServiceMsTeamsChannel 'Microsoft.BotService/botServices/channels@2021-03-01' = {
  parent: botService
  location: 'global'
  name: 'MsTeamsChannel'
  properties: {
    channelName: 'MsTeamsChannel'
  }
}

// Connect the bot service to Azure Active Directory by adding the OAuth connection setting through the Azure Active Directory v2 provider
resource aadV2BotServiceConnection 'Microsoft.BotService/botServices/connections@2021-03-01' = {
  parent: botService
  name: 'aadV2Provider'
  location: 'global'
  properties: {
    serviceProviderDisplayName: 'Azure Active Directory v2'
    serviceProviderId: '30dd229c-58e3-4a48-bdfd-91ec48eb906c'
    scopes: 'User.Read Team.ReadBasic.All Channel.ReadBasic.All ChatMessage.Read ProfilePhoto.Read.All ChannelMessage.Read.All Files.Read.All'
    parameters: [
      {
        key: 'clientId'
        value: aadAppClientId
      }
      {
        key: 'clientSecret'
        value: aadAppClientSecret
      }
      {
        key: 'tenantID'
        value: botType == 'SingleTenant' ? botTenantId : 'common'
      }
      // {
      //   key: 'tokenExchangeUrl'
      //   value: 'api://botid-${aadAppClientId}'
      // }
    ]
  }
}

// Connect the bot service to Azure Active Directory by adding the OAuth connection setting through the Azure Active Directory provider
resource aadBotServiceConnection 'Microsoft.BotService/botServices/connections@2021-03-01' = {
  parent: botService
  name: 'aadProvider'
  location: 'global'
  properties: {
    serviceProviderDisplayName: 'Azure Active Directory'
    serviceProviderId: '5232e24f-b6c6-4920-b09d-d93a520c92e9'
    parameters: [
      {
        key: 'clientId'
        value: aadAppClientId
      }
      {
        key: 'clientSecret'
        value: aadAppClientSecret
      }
      {
        key: 'grantType'
        value: 'authorization_code'
      }
      {
        key: 'loginUri'
        value: botLoginUrl
      }
      {
        key: 'tenantID'
        value: botType == 'SingleTenant' ? botTenantId : 'common'
      }
      {
        key: 'resourceUri'
        value: 'https://graph.microsoft.com'
      }
    ]
  }
}

output BOT_ID string = botAadAppClientId
output BOT_TYPE string = botType
output BOT_TENANT_ID string = botTenantId
output SECRET_BOT_PASSWORD string = aadAppClientSecret
output BOT_CONNECTION_NAME string = aadBotServiceConnection.name

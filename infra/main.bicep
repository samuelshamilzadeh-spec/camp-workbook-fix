// Everything this job needs to run in Azure.
//
//   az deployment group create \
//     --resource-group camp-workbook-sync \
//     --template-file infra/main.bicep \
//     --parameters appName=camp-workbook-sync graphClientId=<app id>
//
// Deliberate choices, each of which costs money or safety to get wrong:
//
//   * **Consumption (Y1), not Flex Consumption.** Flex bills a minimum of
//     1,000 ms per execution at a larger default instance size. At a 5-second
//     timer that is ~17,000 invocations a day, and the difference between the
//     two plans is the entire cost story. Consumption bills 128 MB x 100 ms
//     minimums and this workload fits inside the free monthly grant.
//   * **A user-assigned managed identity, and no client secret.** The Graph
//     permission this job holds is Files.ReadWrite.All — tenant-wide read and
//     write over every OneDrive and SharePoint site. A leaked secret is
//     therefore a tenant-wide incident, so the deployment is built so there is
//     no secret to leak. The role assignment on the identity still has to be
//     granted by an admin; see the output at the bottom.
//   * **A daily cap on Application Insights.** Verbose logging at 17,000
//     invocations a day is the one thing here that can generate a surprising
//     bill, so the cap is set in code rather than left to whoever clicks
//     through the portal.
//   * **HTTPS only, TLS 1.2, no FTP, no public blob access.** This job holds a
//     credential that reaches patient data.

@description('Base name. Resources derive from it, so keep it short and lowercase.')
@minLength(3)
@maxLength(20)
param appName string

@description('Region. Put it near the SharePoint tenant, not near you.')
param location string = resourceGroup().location

@description('The Entra app registration this job authenticates as.')
param graphClientId string

@description('Drive holding the workbook. From `npm run resolve`.')
param graphDriveId string = ''

@description('The workbook driveItem id. From `npm run resolve`.')
param graphItemId string = ''

@description('Phase gate. Nothing above this is applied. See the README phase table.')
@allowed(['0', '1', '2', '3', '4', '5'])
param syncPhase string = '1'

@description('Dry run stays available at every phase. true = plan and log, write nothing.')
@allowed(['true', 'false'])
param syncDryRun string = 'true'

@description('Refuses every write path while false. Flip only after checking LAYOUT.')
@allowed(['true', 'false'])
param syncLayoutVerified string = 'false'

@description('Cap on Application Insights ingestion. The bill lives here.')
param appInsightsDailyCapGb int = 1

// Storage account names are globally unique across all of Azure, so a readable
// one derived from appName is a coin flip on whether the deployment fails at the
// last step. Derived from the resource group id instead: stable across re-runs of
// the same deployment, unique against everyone else's.
var storageName = take('st${uniqueString(resourceGroup().id)}${replace(toLower(appName), '-', '')}', 24)
var identityName = '${appName}-id'
var planName = '${appName}-plan'
var insightsName = '${appName}-ai'

// --- Identity ---------------------------------------------------------------
//
// Created here, but the Graph permission on it is a separate admin grant: Bicep
// cannot assign an app role on Microsoft Graph. The command is in the outputs.
resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: identityName
  location: location
}

// --- Storage ----------------------------------------------------------------
//
// Required by the Functions runtime, and it also holds this job's checkpoint
// blob — the loop guard and the cold-scan cursor.
resource storage 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageName
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    supportsHttpsTrafficOnly: true
    allowBlobPublicAccess: false
    allowSharedKeyAccess: true
  }
}

resource stateContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  name: '${storage.name}/default/camp-sync-state'
  properties: { publicAccess: 'None' }
}

// --- Telemetry --------------------------------------------------------------
resource workspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: '${appName}-logs'
  location: location
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
    workspaceCapping: { dailyQuotaGb: appInsightsDailyCapGb }
  }
}

resource insights 'Microsoft.Insights/components@2020-02-02' = {
  name: insightsName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: workspace.id
    // Sampling is also configured in host.json; both are needed, because
    // host.json governs what the SDK sends and this governs what is kept.
    SamplingPercentage: 20
  }
}

// --- Hosting ----------------------------------------------------------------
resource plan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: planName
  location: location
  // Y1 is Consumption. Read the note at the top before changing this.
  sku: { name: 'Y1', tier: 'Dynamic' }
  properties: { reserved: true }
}

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: appName
  location: location
  kind: 'functionapp,linux'
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: { '${identity.id}': {} }
  }
  properties: {
    serverFarmId: plan.id
    httpsOnly: true
    siteConfig: {
      // Node 20 reached end of life on 2026-04-30. The Azure CLI warns about it
      // on every command and an unsupported runtime stops getting security
      // patches, which matters more than usual for a process holding a
      // tenant-wide Graph credential. 22 is LTS into 2027.
      linuxFxVersion: 'Node|22'
      minTlsVersion: '1.2'
      ftpsState: 'Disabled'
      http20Enabled: true
      appSettings: [
        { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
        { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'node' }
        { name: 'WEBSITE_NODE_DEFAULT_VERSION', value: '~22' }
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storage.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storage.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storage.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storage.listKeys().keys[0].value}'
        }
        { name: 'WEBSITE_CONTENTSHARE', value: toLower(appName) }
        { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: insights.properties.ConnectionString }

        // --- This job's settings, mirroring .env.example -------------------
        { name: 'AZURE_TENANT_ID', value: subscription().tenantId }
        { name: 'AZURE_CLIENT_ID', value: graphClientId }
        { name: 'AZURE_USE_MANAGED_IDENTITY', value: 'true' }
        { name: 'AZURE_MANAGED_IDENTITY_CLIENT_ID', value: identity.properties.clientId }
        { name: 'GRAPH_DRIVE_ID', value: graphDriveId }
        { name: 'GRAPH_ITEM_ID', value: graphItemId }
        { name: 'SYNC_PHASE', value: syncPhase }
        { name: 'SYNC_DRY_RUN', value: syncDryRun }
        { name: 'SYNC_LAYOUT_VERIFIED', value: syncLayoutVerified }
        { name: 'SYNC_STATE_CONTAINER', value: 'camp-sync-state' }
        { name: 'SYNC_READ_CONCURRENCY', value: '8' }
        { name: 'SYNC_MAX_SHEETS_PER_CYCLE', value: '90' }

        // The rollback lever. Flipping this to true stops the timer without a
        // deploy and without the Function App being healthy. Everyone involved
        // in a cutover should know it exists.
        { name: 'AzureWebJobs.syncTimer.Disabled', value: 'false' }
      ]
    }
  }
  dependsOn: [stateContainer]
}

// The identity needs to read and write its own checkpoint blobs.
resource blobContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storage.id, identity.id, 'ba92f5b4-2d11-453d-a403-e96b0029c9fe')
  scope: storage
  properties: {
    principalId: identity.properties.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'ba92f5b4-2d11-453d-a403-e96b0029c9fe' // Storage Blob Data Contributor
    )
  }
}

output functionAppName string = functionApp.name
output identityPrincipalId string = identity.properties.principalId
output identityClientId string = identity.properties.clientId

@description('THE step that makes this work. Bicep cannot assign a Microsoft Graph app role, so run this as an admin.')
output grantGraphToManagedIdentity string = 'az rest --method POST --uri "https://graph.microsoft.com/v1.0/servicePrincipals/${identity.properties.principalId}/appRoleAssignments" --headers "Content-Type=application/json" --body \'{"principalId":"${identity.properties.principalId}","resourceId":"GRAPH_SP_OBJECT_ID","appRoleId":"75359482-378d-4052-8f01-80520e7db3cd"}\'  # get GRAPH_SP_OBJECT_ID from: az ad sp show --id 00000003-0000-0000-c000-000000000000 --query id -o tsv'

@description('Only needed for the CLI scripts, which authenticate as the app registration rather than as the identity. Probably already granted.')
output grantGraphToAppRegistration string = 'az ad app permission add --id ${graphClientId} --api 00000003-0000-0000-c000-000000000000 --api-permissions 75359482-378d-4052-8f01-80520e7db3cd=Role && az ad app permission admin-consent --id ${graphClientId}'

output storageAccountName string = storage.name

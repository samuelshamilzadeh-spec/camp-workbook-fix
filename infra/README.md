# Provisioning

None of this exists yet. Every result this project has produced came from
running the code by hand against the live workbook; nothing runs on a timer.

`main.bicep` creates the lot: resource group contents, storage (which also holds
this job's checkpoint blob), Log Analytics with a daily cap, Application
Insights, a Consumption plan, the Function App, and a user-assigned managed
identity with Storage Blob Data Contributor on its own state container.

## Order

Run this **from the repo directory**, not from your home directory — the
template path is relative.

**Single quotes around the drive id.** It contains a `!`, and inside double
quotes bash treats that as history expansion and fails with
`event not found` before Azure ever sees the command.

```bash
cd /path/to/camp-workbook-fix

RG=camp-workbook-sync
APP=camp-workbook-sync                              # must be globally unique
CLIENT_ID=27e61a22-9b04-4124-a5e0-5b691f0435c4      # the Entra app registration

az group create --name "$RG" --location eastus

az deployment group create \
  --resource-group "$RG" \
  --template-file infra/main.bicep \
  --parameters appName="$APP" \
               graphClientId="$CLIENT_ID" \
               graphDriveId='<from npm run resolve>' \
               graphItemId='<from npm run resolve>'
```

`appName` becomes `appName.azurewebsites.net`, so it has to be unique across
Azure. If the deployment fails on the site name, pick another — nothing else
depends on it. The storage account name is derived from the resource group id
rather than from `appName`, precisely so that one cannot collide.

## Do you have a Global Admin awake?

If not, **skip the next section entirely** and use the client secret. The
managed identity is the better long-term answer and it is what this template
deploys, but it needs an app role assignment on Microsoft Graph that only a
Global Administrator or Privileged Role Administrator can make. Discovering that
after the resources are up, at four in the morning, is a bad time to discover it.

The app registration already holds `Files.ReadWrite.All` with admin consent —
it is what every CLI script in this repo authenticates with — so the Function
App can use the same credential today and be moved to the identity later
without a redeploy:

```bash
az functionapp config appsettings set -g "$RG" -n "$APP" \
  --settings AZURE_USE_MANAGED_IDENTITY=false AZURE_CLIENT_SECRET='<the secret>'
```

That is no worse than the status quo — the same credential, the same scope, in
a place designed to hold it. Moving to the identity afterwards is:

```bash
az functionapp config appsettings set -g "$RG" -n "$APP" --settings AZURE_USE_MANAGED_IDENTITY=true
az functionapp config appsettings delete -g "$RG" -n "$APP" --setting-names AZURE_CLIENT_SECRET
```

## The grant, for when an admin is available

The Function App authenticates as its **managed identity**, not as the app
registration. That distinction is easy to miss and it is the whole difference
between a working deployment and one that starts up, reads nothing, and logs
403s: the app registration already holds `Files.ReadWrite.All` — that is what
every CLI script in this repo uses — and the managed identity holds nothing at
all until somebody grants it.

Bicep cannot assign an app role on Microsoft Graph, so this is a manual step for
someone with Global Administrator or Privileged Role Administrator:

```bash
PRINCIPAL_ID=<identityPrincipalId from the deployment outputs>
GRAPH_SP=$(az ad sp show --id 00000003-0000-0000-c000-000000000000 --query id -o tsv)

az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PRINCIPAL_ID/appRoleAssignments" \
  --headers 'Content-Type=application/json' \
  --body "{\"principalId\":\"$PRINCIPAL_ID\",\"resourceId\":\"$GRAPH_SP\",\"appRoleId\":\"75359482-378d-4052-8f01-80520e7db3cd\"}"
```

`75359482-378d-4052-8f01-80520e7db3cd` is `Files.ReadWrite.All`. Confirm it took:

```bash
az rest --method GET \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PRINCIPAL_ID/appRoleAssignments" \
  --query "value[].appRoleId" -o tsv
```

The `grantGraphToAppRegistration` output is only for the CLI scripts, which
authenticate with the client secret. If `npm run reconcile` already works, that
grant exists and you do not need it.

## Then, and not before

The template deploys with `syncPhase=1`, `syncDryRun=true` and
`syncLayoutVerified=false`, which is a job that reads and logs and cannot write
anything. That is on purpose. Moving it forward is three app settings and no
deploy:

```bash
az functionapp config appsettings set -g "$RG" -n "$APP" --settings SYNC_PHASE=2
az functionapp config appsettings set -g "$RG" -n "$APP" --settings SYNC_LAYOUT_VERIFIED=true
az functionapp config appsettings set -g "$RG" -n "$APP" --settings SYNC_DRY_RUN=false
```

Watch a cycle before flipping the last one. `cycle.plan` says what it would do;
`cycle.applied` says what it did.

## The rollback lever

```bash
az functionapp config appsettings set -g "$RG" -n "$APP" \
  --settings AzureWebJobs.syncTimer.Disabled=true
```

It takes effect without a deploy and does not depend on the Function App being
healthy. Everyone involved in a cutover should know it exists before they need
it.

## No client secret

`AZURE_USE_MANAGED_IDENTITY=true` and no `AZURE_CLIENT_SECRET` app setting. The
Graph permission this job holds is `Files.ReadWrite.All` — tenant-wide read and
write across every OneDrive and SharePoint site, not just this workbook. A leaked
secret is a tenant-wide incident, so the deployment is built so there is no
secret to leak.

Secrets used during development should be deleted once this is running:

```bash
az ad app credential list --id "$CLIENT_ID" -o table
az ad app credential delete --id "$CLIENT_ID" --key-id <keyId>
```

## Costs, and the one that bites

At a 5-second timer this is ~17,000 invocations a day. On Consumption that fits
inside the free monthly grant; on Flex Consumption it does not, because Flex
bills a minimum of 1,000 ms per execution at a larger instance size. The plan SKU
is the whole cost story and the template pins `Y1`.

The other one is telemetry. Verbose logging at that invocation rate is what
generates a surprising bill, so the workspace carries a daily cap
(`appInsightsDailyCapGb`, default 1 GB) and `host.json` samples at 2 items/second
with exceptions excluded. Both are needed: `host.json` governs what the SDK
sends, the cap governs what is kept.

## GitHub Actions

`.github/workflows/deploy.yml` signs in with OIDC, so there is no publish profile
in the repo secrets. It needs a federated credential on the app registration for
this repo, plus repo secrets `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`,
`AZURE_SUBSCRIPTION_ID` and the variable `AZURE_FUNCTIONAPP_NAME`.

```bash
az ad app federated-credential create --id "$CLIENT_ID" --parameters '{
  "name": "github-main",
  "issuer": "https://token.actions.githubusercontent.com",
  "subject": "repo:samuelshamilzadeh-spec/camp-workbook-fix:ref:refs/heads/main",
  "audiences": ["api://AzureADTokenExchange"]
}'
```

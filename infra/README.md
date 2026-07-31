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

## Deploying the code

**Do not use `az functionapp deployment source config-zip` here.** It reports
success, sets `WEBSITE_RUN_FROM_PACKAGE=1`, and leaves you with a Function App
that lists zero functions. Linux Consumption does not support the `1` form of
that setting — it only accepts a URL — so the host mounts no content and finds
nothing to run. Every other diagnostic looks healthy: the zip is correct, the
app settings are correct, the deployment succeeded.

Use Core Tools, which handles this correctly and is present in Cloud Shell:

```bash
npm ci && npm run build
func azure functionapp publish "$APP"
```

If `func` is unavailable, do what it does — put the package in blob storage and
point the setting at it:

```bash
STORAGE=$(az functionapp config appsettings list -g "$RG" -n "$APP" \
  --query "[?name=='WEBSITE_CONTENTSHARE']" -o tsv >/dev/null; \
  az storage account list -g "$RG" --query "[0].name" -o tsv)
KEY=$(az storage account keys list -g "$RG" -n "$STORAGE" --query "[0].value" -o tsv)

npm ci && npm run build && npm prune --omit=dev
rm -f ../app.zip && zip -rq ../app.zip dist node_modules host.json package.json && npm ci

az storage container create --account-name "$STORAGE" --account-key "$KEY" --name deployments
az storage blob upload --account-name "$STORAGE" --account-key "$KEY" \
  --container-name deployments --name app.zip --file ../app.zip --overwrite

EXPIRY=$(date -u -d '+2 years' '+%Y-%m-%dT%H:%MZ')
SAS=$(az storage blob generate-sas --account-name "$STORAGE" --account-key "$KEY" \
  --container-name deployments --name app.zip --permissions r --expiry "$EXPIRY" -o tsv)

az functionapp config appsettings set -g "$RG" -n "$APP" \
  --settings WEBSITE_RUN_FROM_PACKAGE="https://$STORAGE.blob.core.windows.net/deployments/app.zip?$SAS"
az functionapp restart -g "$RG" -n "$APP"
```

The SAS expiry is a real deadline: when it passes the app stops being able to
read its own code. Core Tools manages this for you, which is the better reason
to prefer it.

Either way, confirm before moving on:

```bash
sleep 30 && az functionapp function list -g "$RG" -n "$APP" -o table
```

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

### Reading the logs

**`az webapp log tail` does not work here.** A Linux Consumption Function App
has no SCM log stream, so it returns a 404 that looks like a broken deployment
and is not one. Logs go to Application Insights.

First, check the deployed package actually registered the function — this is the
thing that most often goes wrong after a zip deploy, and it fails silently:

```bash
az functionapp function list -g "$RG" -n "$APP" -o table
```

`syncTimer` should be listed. An empty table is the failure to expect, and it
has three usual causes, in order of likelihood:

1. **`AzureWebJobsFeatureFlags` is not set to `EnableWorkerIndexing`.** The v4
   Node model registers functions from code, not from a `function.json` beside
   each one. Without the flag the host looks for those files, finds none, and
   reports zero functions — a deployment that succeeded and does nothing.

   ```bash
   az functionapp config appsettings set -g "$RG" -n "$APP" \
     --settings AzureWebJobsFeatureFlags=EnableWorkerIndexing
   az functionapp restart -g "$RG" -n "$APP"
   ```

2. **`host.json` is not at the zip's root.** `zip -r ../app.zip dist node_modules
   host.json package.json` from the repo root gets this right; zipping the
   containing folder does not. Check with `unzip -l ../app.zip | head`.

3. **`dist/` was never built**, so there is nothing for `main` in package.json to
   match. `npm run build` first, and confirm `dist/src/functions/syncTimer.js`
   exists before zipping.

Note that the app takes a few seconds to re-index after a restart, so re-run the
list rather than concluding from the first attempt.

Then read what it is doing. Allow two or three minutes for ingestion:

**In Azure Cloud Shell this fails**: its managed identity cannot get a token for
`api.applicationinsights.io`. Either run it from a workstation with
`az login`, or read the same data in the portal under the Function App's
*Monitor* blade, which needs no extra token.

```bash
az extension add --name application-insights --only-show-errors
APPID=$(az monitor app-insights component show -g "$RG" -a "$APP-ai" --query appId -o tsv)

az monitor app-insights query --app "$APPID" --analytics-query \
  "traces | where timestamp > ago(30m) | order by timestamp desc | take 30 | project timestamp, message" \
  --query "tables[0].rows" -o tsv
```

`cycle.plan` says what it would do; `cycle.applied` says what it did. All zeroes
in `cycle.plan` means it agrees with `npm run reconcile` run from a laptop.

Exceptions are excluded from sampling, so a failure is never dropped:

```bash
az monitor app-insights query --app "$APPID" --analytics-query \
  "exceptions | where timestamp > ago(30m) | project timestamp, outerMessage" \
  --query "tables[0].rows" -o tsv
```

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

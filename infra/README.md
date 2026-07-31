# Provisioning

None of this exists yet. Every result this project has produced came from
running the code by hand against the live workbook; nothing runs on a timer.

`main.bicep` creates the lot: resource group contents, storage (which also holds
this job's checkpoint blob), Log Analytics with a daily cap, Application
Insights, a Consumption plan, the Function App, and a user-assigned managed
identity with Storage Blob Data Contributor on its own state container.

## Order

```bash
RG=camp-workbook-sync
APP=camp-workbook-sync
CLIENT_ID=27e61a22-9b04-4124-a5e0-5b691f0435c4      # the Entra app registration

az group create --name "$RG" --location eastus

az deployment group create \
  --resource-group "$RG" \
  --template-file infra/main.bicep \
  --parameters appName="$APP" graphClientId="$CLIENT_ID" \
               graphDriveId="<from npm run resolve>" \
               graphItemId="<from npm run resolve>"
```

The deployment prints a `grantGraphPermission` command. **Run it as an admin** —
Bicep cannot assign an app role on Microsoft Graph, so the identity has no access
to the workbook until somebody does.

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

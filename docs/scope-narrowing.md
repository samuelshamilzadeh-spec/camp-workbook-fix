# Narrowing the Graph permission

This deployment uses `Files.ReadWrite.All`, chosen over the `Sites.Selected`
the build brief specified. See "Scope decision, recorded" in the README for the
reasoning.

If the scope is ever revisited — a security review, an audit, a BAA renewal, or
simply someone with an hour to spare — this is the whole path back. **No code
changes.** The Graph calls this project makes are identical under either
permission.

Budget: about fifteen minutes, most of it waiting on an admin.

---

## Steps

1. **Grant the app on the one site.** Needs `Sites.FullControl.All`, or a
   SharePoint admin running PnP PowerShell:

   ```powershell
   Connect-PnPOnline -Url https://premierassist.sharepoint.com/sites/<site> -Interactive
   Grant-PnPAzureADAppSitePermission `
     -AppId 27e61a22-9b04-4124-a5e0-5b691f0435c4 `
     -DisplayName "camp-workbook-sync" `
     -Permissions Write
   ```

   Or the raw Graph call:

   ```http
   POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
   {
     "roles": ["write"],
     "grantedToIdentities": [
       { "application": { "id": "{appClientId}", "displayName": "camp-workbook-sync" } }
     ]
   }
   ```

2. **Add `Sites.Selected` and consent:**

   ```bash
   az ad sp show --id 00000003-0000-0000-c000-000000000000 \
     --query "appRoles[?value=='Sites.Selected'].{id:id,value:value}" -o table

   az ad app permission add --id {appClientId} \
     --api 00000003-0000-0000-c000-000000000000 \
     --api-permissions {roleId}=Role

   az ad app permission admin-consent --id {appClientId}
   ```

3. **Verify the narrow permission works before removing the broad one.** Run
   `npm run inspect` — it is read-only and exercises the same auth path as the
   Function. Both permissions are live at this point, so a failure here means
   the site grant did not take.

4. **Remove `Files.ReadWrite.All`:**

   ```bash
   az ad app permission delete --id {appClientId} \
     --api 00000003-0000-0000-c000-000000000000 \
     --api-permissions 75359482-378d-4052-8f01-80520e7db3cd
   ```

   Confirm the role id with the same `az ad sp show` query first. Removing an
   app role assignment can take a few minutes to propagate, and cached tokens
   live up to an hour — the running Function may keep working briefly on an old
   token. That is expected, not a failed removal.

5. **Confirm the scoping actually bit.** With `Sites.Selected` alone, a request
   against any *other* site must return 403. That negative test is the only
   real proof the narrowing worked; step 3 passing proves nothing about what is
   now denied.

6. **Update the README.** Delete the "Scope decision, recorded" block and put
   the `Sites.Selected` description back. Stale security documentation is worse
   than none.

---

## Do this at the same time

Whenever the scope is revisited, these two are cheap and address the same risk:

- **Drop the client secret** in favour of a user-assigned managed identity
  (`AZURE_USE_MANAGED_IDENTITY=true`). Under a tenant-wide permission this is
  the single highest-value change available, because it removes the thing that
  can leak.
- **Split the deploy identity.** If the GitHub Actions federated credential was
  put on this same app registration, move it to a separate one with Contributor
  on the resource group and no Graph permissions at all. Right now a compromised
  CI pipeline reaches the workbook.

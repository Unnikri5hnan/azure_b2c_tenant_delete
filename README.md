# b2c-tenant-delete.sh — one script, one invocation

Everything from the previous multi-file version (`setup-app.sh`,
`get_graph_token.py`, `delete-b2c-tenant.sh`, `run-all.sh`) is now in this
single file. Run `./b2c-tenant-delete.sh --help` for full usage — the
header comment in the script is the primary documentation and stays in
sync with the code. This file just covers the essentials.

## Quick start

```bash
chmod +x b2c-tenant-delete.sh
./b2c-tenant-delete.sh \
  --b2c-tenant contoso1.onmicrosoft.com \
  --subscription 00000000-0000-0000-0000-000000000001 \
  --resource-group rg-b2c-contoso1 \
  --dry-run
```

That's it — no separate setup step. On first run it will:
1. Check for `az`, `jq`, `curl`, `python3`, and the `msal` Python package,
   installing anything missing that it can (apt-get / pip3 --user).
2. Sign you in once (`az login`) to create its own small automation app in
   the tenant (needed because Azure CLI's own app is blocked from getting
   the Graph permissions this requires — see the script header for why).
3. Sign you in once more via device code through that app, then use that
   single session for all cleanup steps AND the final tenant deletion.

Drop `--dry-run` once the output looks right. If you want to empty the
tenant but keep it around for now (not delete it yet), add `--cleanup-only`
instead — see "Cleaning up without deleting the tenant" below.

## Re-running

The automation app's id and the device-code session are both cached to
`~/.cache/b2c-cleanup/`. Re-running against the same tenant normally needs
**zero** interactive prompts — it silently reuses both. Use `--force-setup`
to redo the app/permissions from scratch, or `--no-cache` to force a fresh
device-code sign-in.

## Many tenants

```bash
cp tenants.example.csv tenants.csv
# edit with your real values (mgmt_tenant / app_client_id / mgmt_app_client_id may be left blank)
./b2c-tenant-delete.sh --csv tenants.csv --dry-run
./b2c-tenant-delete.sh --csv tenants.csv --yes
```

Keeps going if one tenant fails, and writes `results-<timestamp>.log`.

## You'll also need, separately (this script can't grant these)

- **Global Administrator** in the B2C tenant.
- **Owner or Contributor (Azure RBAC)** on the resource group/subscription
  holding the tenant's `b2cDirectories` resource — if the final delete
  fails with `AuthorizationFailed`, this is what's missing:
  ```bash
  az role assignment create --assignee <your-upn> --role Contributor --resource-group <your-rg>
  ```

## Cleaning up without deleting the tenant

If you just want the tenant emptied (users, app registrations, enterprise
applications, identity providers, user flows, IEF keys/policies, and the
automation app itself removed) but want to keep the tenant shell around for
now, add `--cleanup-only`:

```bash
./b2c-tenant-delete.sh \
  --b2c-tenant contoso1.onmicrosoft.com \
  --subscription 00000000-0000-0000-0000-000000000001 \
  --resource-group rg-b2c-contoso1 \
  --cleanup-only
```

This runs steps 1-8 and stops — step 9 (the actual ARM tenant delete) is
skipped, and the tenant is left in place. Drop the flag later and re-run
when you're ready to delete it for real.

## Scoping enterprise-application deletion

By default, step 3 deletes **all** enterprise applications / service
principals — this matches what Microsoft's own "Delete tenant" page in the
portal requires. To scope it (e.g. for testing), use:
```bash
--sp-name-filter "Microsoft Graph Command Line Tools"
```
Note this will likely leave the tenant-delete checklist still red on that
category, since Microsoft requires all of them gone — the script warns you
how many are being left behind.

#!/usr/bin/env bash
#
# b2c-tenant-delete.sh
#
# ONE SCRIPT, ONE INVOCATION: cleans up and deletes an Azure AD B2C tenant
# (or a whole CSV list of them), following Microsoft's own "Delete tenant"
# checklist:
#   https://learn.microsoft.com/azure/active-directory-b2c/tutorial-delete-tenant
#
# WHAT IT DOES, PER TENANT
#   0. Checks/installs prerequisites (az, jq, curl, python3, the `msal`
#      Python package) - see PREREQUISITES below.
#   1. Ensures a small "automation app" exists in the tenant with the
#      Graph + ARM permissions Azure CLI's own app is blocked from getting
#      (see AUTH MODEL). Reuses one from a previous run if already cached -
#      no `az login` needed at all in that case.
#   2. Deletes all users (removing directory-role membership first if a
#      user - commonly a guest/admin account - holds one)
#   3. Deletes all app registrations (incl. b2c-extensions-app / IEF apps)
#   4. Deletes all enterprise applications / service principals - a
#      DIFFERENT object type from app registrations, also required empty
#      by Microsoft's checklist (commonly includes an auto-provisioned
#      "Microsoft Graph Command Line Tools" entry)
#   5. Deletes all identity providers
#   6. Deletes all user flows
#   7. Deletes all Identity Experience Framework (IEF) policy keys
#   8. Deletes all IEF custom policies
#   9. Deletes the automation app itself, now that it's no longer needed
#  10. Deletes the tenant itself (async ARM call on
#      Microsoft.AzureActiveDirectory/b2cDirectories, polled until gone)
#
# Any failure in steps 2-9 stops the script BEFORE it touches the tenant,
# so you never attempt to delete a tenant that still has resources.
#
# LOGIN MODEL (minimal prompts)
# ------------------------------
# First run ever, for a given tenant: up to two sign-ins -
#   (a) `az login` to create/verify the automation app (needs an existing
#       Global Admin session), and
#   (b) one device-code sign-in through that app, used for EVERYTHING else
#       (Graph cleanup steps 2-8 AND the final ARM tenant delete in step
#       10) - because a delegated ARM token is authorized against YOUR
#       Azure RBAC roles regardless of which app requested it, there's no
#       need for a second, `az`-specific sign-in just to reach ARM.
# Every run after that, for the same tenant: the automation app's id is
# cached to disk, so step (a) - and its `az login` - is skipped entirely.
# The device-code session is also cached and silently refreshed, so
# normally you won't see any interactive prompt at all on a retry.
# Use --force-setup to bypass the app cache (e.g. permissions changed) and
# --no-cache to bypass the device-code session cache.
#
# A second, unavoidable sign-in only happens if --mgmt-tenant is a
# genuinely different Azure AD tenant from --b2c-tenant (common when the
# subscription lives in your org's regular tenant, not the B2C tenant) -
# the script sets up and caches an automation app there too.
#
# AUTH MODEL (why an automation app is needed at all)
# -----------------------------------------------------
# A plain `az login` token cannot manage identity providers, user flows, or
# IEF policies/keys - you'll see:
#   "The application does not have any of the required delegated permissions
#    (IdentityUserFlow.Read.All, IdentityUserFlow.ReadWrite.All)..."
# This is because Azure CLI's own app registration is a Microsoft
# first-party app, and Microsoft does not allow tenants to grant it extra
# scopes (AADSTS65002) - being Global Admin doesn't get around this. The
# fix is a small app YOU own, self-consented, which this script manages
# for you automatically.
#
# PREREQUISITES
# --------------
# Needs: az, jq, curl, python3 (+ pip). The script checks each of these at
# startup and attempts to install anything missing (apt-get if available
# and root/sudo works, or pip3 --user for the Python `msal` package). If it
# can't auto-install something, it tells you exactly what to install and
# where from, then exits.
#
# You ALSO need, separately (this script cannot grant you these):
#   - Global Administrator in the B2C tenant
#   - Owner or Contributor (Azure RBAC) on the resource group/subscription
#     that holds the tenant's b2cDirectories resource - if the final ARM
#     delete fails with AuthorizationFailed, this is what's missing.
#
# USAGE
# -----
#   Single tenant:
#     ./b2c-tenant-delete.sh \
#         --b2c-tenant <tenantIdOrDomain> \
#         --subscription <subscriptionId> \
#         --resource-group <resourceGroupName> \
#         [--mgmt-tenant <mgmtTenantId>]        # defaults to --b2c-tenant
#         [--app-client-id <clientId>]          # skip auto-setup, use this app directly
#         [--mgmt-app-client-id <clientId>]     # same, for the mgmt tenant, if it differs
#         [--app-name <name>]                   # automation app display name (default: b2c-cleanup-automation)
#         [--sp-name-filter "Name1,Name2"]      # step 4 only: only touch these enterprise apps (see README notes below)
#         [--force-setup]                       # ignore cached automation-app id, redo setup
#         [--no-cache]                          # ignore cached device-code session, force fresh sign-in
#         [--dry-run]                           # list what would be deleted, delete nothing
#         [--yes]                               # skip the interactive confirmation prompt
#
#   Many tenants from a CSV (no header row):
#     b2c_tenant,subscription_id,resource_group,mgmt_tenant,app_client_id,mgmt_app_client_id
#     (mgmt_tenant / app_client_id / mgmt_app_client_id may be left blank)
#
#     ./b2c-tenant-delete.sh --csv tenants.csv [--dry-run] [--yes]
#
# EXIT CODES
#   0 success, 1 bad args / preflight / setup failure, 2 cleanup failure,
#   3 tenant-delete failure. In --csv mode, the script keeps going after a
#   failed tenant and reports a summary at the end.
#
set -u -o pipefail

# ============================================================================
# Globals / defaults
# ============================================================================
DRY_RUN=false
AUTO_YES=false
FORCE_SETUP=false
NO_CACHE=false
APP_NAME="b2c-cleanup-automation"
SP_NAME_FILTER=""
CSV_FILE=""
GRAPH_API="https://graph.microsoft.com"
ARM_API="https://management.azure.com"
GRAPH_APP_ID="00000003-0000-0000-c000-000000000000"   # Microsoft Graph's own appId, constant everywhere
ARM_APP_ID="797f4846-ba00-4fd7-ba43-dac1f8f63013"      # Azure Service Management API's own appId, constant everywhere
CACHE_DIR="${HOME}/.cache/b2c-cleanup"
LOG_PREFIX=""

log()  { echo "[$(date '+%Y-%m-%d %H:%M:%S')]${LOG_PREFIX} $*"; }
err()  { echo "[$(date '+%Y-%m-%d %H:%M:%S')]${LOG_PREFIX} ERROR: $*" >&2; }
die()  { err "$*"; exit "${2:-1}"; }

# ============================================================================
# Arg parsing
# ============================================================================
B2C_TENANT=""; SUBSCRIPTION_ID=""; RESOURCE_GROUP=""; MGMT_TENANT=""
APP_CLIENT_ID=""; MGMT_APP_CLIENT_ID=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --b2c-tenant) B2C_TENANT="$2"; shift 2;;
    --subscription) SUBSCRIPTION_ID="$2"; shift 2;;
    --resource-group) RESOURCE_GROUP="$2"; shift 2;;
    --mgmt-tenant) MGMT_TENANT="$2"; shift 2;;
    --app-client-id) APP_CLIENT_ID="$2"; shift 2;;
    --mgmt-app-client-id) MGMT_APP_CLIENT_ID="$2"; shift 2;;
    --app-name) APP_NAME="$2"; shift 2;;
    --sp-name-filter) SP_NAME_FILTER="$2"; shift 2;;
    --csv) CSV_FILE="$2"; shift 2;;
    --force-setup) FORCE_SETUP=true; shift;;
    --no-cache) NO_CACHE=true; shift;;
    --dry-run) DRY_RUN=true; shift;;
    --yes) AUTO_YES=true; shift;;
    -h|--help) grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0;;
    *) die "Unknown argument: $1";;
  esac
done

if [[ -z "$CSV_FILE" ]]; then
  : "${B2C_TENANT:?--b2c-tenant is required (or use --csv)}"
  : "${SUBSCRIPTION_ID:?--subscription is required (or use --csv)}"
  : "${RESOURCE_GROUP:?--resource-group is required (or use --csv)}"
fi

mkdir -p "$CACHE_DIR"

# ============================================================================
# Preflight: check + attempt to install prerequisites
# ============================================================================
_apt_install() {
  local pkg="$1"
  if ! command -v apt-get >/dev/null 2>&1; then
    return 1
  fi
  if [[ "${EUID:-$(id -u)}" -eq 0 ]]; then
    apt-get update -y >/dev/null 2>&1 && apt-get install -y "$pkg" >/dev/null 2>&1
  elif command -v sudo >/dev/null 2>&1; then
    sudo apt-get update -y >/dev/null 2>&1 && sudo apt-get install -y "$pkg" >/dev/null 2>&1
  else
    return 1
  fi
}

preflight() {
  log "Checking prerequisites ..."

  if ! command -v curl >/dev/null 2>&1; then
    log "curl not found - attempting install ..."
    _apt_install curl || die "curl is required and could not be auto-installed. Install it manually and re-run."
  fi

  if ! command -v jq >/dev/null 2>&1; then
    log "jq not found - attempting install ..."
    _apt_install jq || die "jq is required and could not be auto-installed. Install it manually (e.g. 'apt-get install jq') and re-run."
  fi

  if ! command -v python3 >/dev/null 2>&1; then
    log "python3 not found - attempting install ..."
    _apt_install python3 || die "python3 is required and could not be auto-installed. Install it manually and re-run."
  fi

  if ! command -v pip3 >/dev/null 2>&1; then
    log "pip3 not found - attempting install ..."
    _apt_install python3-pip || true   # not fatal yet; msal check below will catch it
  fi

  if ! command -v az >/dev/null 2>&1; then
    log "Azure CLI (az) not found - attempting install via Microsoft's install script ..."
    if command -v sudo >/dev/null 2>&1; then
      curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash >/dev/null 2>&1
    elif [[ "${EUID:-$(id -u)}" -eq 0 ]]; then
      curl -sL https://aka.ms/InstallAzureCLIDeb | bash >/dev/null 2>&1
    fi
    command -v az >/dev/null 2>&1 || die "Azure CLI (az) is required and could not be auto-installed. See https://learn.microsoft.com/cli/azure/install-azure-cli and re-run."
  fi

  if ! python3 -c "import msal" >/dev/null 2>&1; then
    log "Python 'msal' package not found - installing ..."
    pip3 install --user msal --quiet 2>/dev/null \
      || pip3 install --user msal --break-system-packages --quiet 2>/dev/null \
      || die "Could not install the 'msal' Python package automatically. Run: pip3 install --user msal"
  fi

  log "All prerequisites present (az, jq, curl, python3, msal)."
}

# ============================================================================
# Tenant GUID resolution + az login reuse
# ============================================================================
resolve_tenant_guid() {
  local input="$1"
  if [[ "$input" =~ ^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$ ]]; then
    echo "${input,,}"
    return 0
  fi
  local issuer guid
  issuer=$(curl -sS "https://login.microsoftonline.com/$input/v2.0/.well-known/openid-configuration" 2>/dev/null | jq -r '.issuer // empty')
  guid=$(grep -oE '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}' <<<"$issuer" | head -n1)
  if [[ -n "$guid" ]]; then echo "${guid,,}"; else echo "$input"; fi
}

ensure_az_login() {
  local tenant="$1"
  local current current_guid target_guid
  current=$(az account show --query tenantId -o tsv 2>/dev/null)
  if [[ -n "$current" ]]; then
    current_guid=$(resolve_tenant_guid "$current")
    target_guid=$(resolve_tenant_guid "$tenant")
    if [[ "$current_guid" == "$target_guid" ]]; then
      log "Already signed in to tenant $tenant (reusing existing az session)."
      return 0
    fi
  fi
  log "Signing in to tenant $tenant ..."
  az login --tenant "$tenant" --allow-no-subscriptions -o none
}

# ============================================================================
# Embedded MSAL device-code token helper (written to a temp file at runtime)
# ============================================================================
TOKEN_HELPER=""
write_token_helper() {
  TOKEN_HELPER=$(mktemp /tmp/.b2c_token_helper.XXXXXX.py)
  cat > "$TOKEN_HELPER" <<'PYEOF'
import sys, os, argparse, json
try:
    import msal
except ImportError:
    print("The 'msal' package is required. Install it with: pip3 install --user msal", file=sys.stderr)
    sys.exit(1)

def default_cache_path(tenant, client_id):
    cache_dir = os.path.join(os.path.expanduser("~"), ".cache", "b2c-cleanup")
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, f"tokencache-{tenant}-{client_id}.bin")

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--client-id", required=True)
    p.add_argument("--tenant", required=True)
    p.add_argument("--scope", required=True)
    p.add_argument("--cache-file", default=None)
    p.add_argument("--no-cache", action="store_true")
    args = p.parse_args()

    cache_path = args.cache_file or default_cache_path(args.tenant, args.client_id)
    cache = msal.SerializableTokenCache()
    if not args.no_cache and os.path.exists(cache_path):
        try:
            with open(cache_path, "r") as f:
                cache.deserialize(f.read())
        except Exception as e:
            print(f"Warning: could not read token cache ({e}); starting fresh.", file=sys.stderr)

    authority = f"https://login.microsoftonline.com/{args.tenant}"
    app = msal.PublicClientApplication(args.client_id, authority=authority, token_cache=cache)
    scopes = args.scope.split()

    result = None
    accounts = app.get_accounts()
    if accounts:
        print(f"Found cached session for {accounts[0].get('username','unknown user')}; trying silent refresh ...", file=sys.stderr)
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print(f"Failed to start device flow: {json.dumps(flow)}", file=sys.stderr)
            sys.exit(1)
        print(flow["message"], file=sys.stderr)
        result = app.acquire_token_by_device_flow(flow)
    else:
        print("Reused cached session - no device code needed.", file=sys.stderr)

    if not args.no_cache and cache.has_state_changed:
        try:
            with open(cache_path, "w") as f:
                f.write(cache.serialize())
            os.chmod(cache_path, 0o600)
        except Exception as e:
            print(f"Warning: could not write token cache ({e})", file=sys.stderr)

    if "access_token" in result:
        print(result["access_token"])
    else:
        print(f"Failed to acquire token: {json.dumps(result)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
PYEOF
}

get_token() {
  local client_id="$1" tenant="$2" scope="$3"
  local extra=()
  $NO_CACHE && extra+=(--no-cache)
  python3 "$TOKEN_HELPER" --client-id "$client_id" --tenant "$tenant" --scope "$scope" "${extra[@]}"
}

cleanup_temp() { [[ -n "$TOKEN_HELPER" ]] && rm -f "$TOKEN_HELPER"; }
trap cleanup_temp EXIT

# ============================================================================
# Automation app: find-or-create, permission, consent - cached across runs
# ============================================================================
marker_file_for_tenant() { echo "$CACHE_DIR/app-$(resolve_tenant_guid "$1").json"; }

ensure_automation_app() {
  local tenant="$1"
  local marker; marker=$(marker_file_for_tenant "$tenant")

  if ! $FORCE_SETUP && [[ -f "$marker" ]]; then
    local cached_id
    cached_id=$(jq -r '.clientId // empty' "$marker" 2>/dev/null)
    if [[ -n "$cached_id" ]]; then
      log "Reusing previously set-up automation app for $tenant: $cached_id (az login skipped)"
      echo "$cached_id"
      return 0
    fi
  fi

  log "Setting up automation app in tenant $tenant (one-time per tenant; needs az login) ..." >&2
  ensure_az_login "$tenant" >&2 || die "Login to $tenant failed (needed to set up the automation app)" 1

  scope_id() { az ad sp show --id "$1" --query "oauth2PermissionScopes[?value=='$2'].id | [0]" -o tsv; }

  local names=(IdentityUserFlow.ReadWrite.All Policy.ReadWrite.TrustFramework IdentityProvider.ReadWrite.All \
               TrustFrameworkKeySet.ReadWrite.All Application.ReadWrite.All User.ReadWrite.All \
               RoleManagement.ReadWrite.Directory)
  local perm_args=()
  local name id
  for name in "${names[@]}"; do
    id=$(scope_id "$GRAPH_APP_ID" "$name")
    [[ -z "$id" || "$id" == "None" ]] && die "Could not resolve permission id for $name"
    perm_args+=("$id=Scope")
  done
  local arm_scope_id
  arm_scope_id=$(scope_id "$ARM_APP_ID" "user_impersonation")
  [[ -z "$arm_scope_id" || "$arm_scope_id" == "None" ]] && die "Could not resolve ARM user_impersonation scope id"

  local existing_id app_id
  existing_id=$(az ad app list --display-name "$APP_NAME" --query "[0].appId" -o tsv 2>/dev/null)
  if [[ -n "$existing_id" && "$existing_id" != "None" ]]; then
    app_id="$existing_id"
    log "Reusing existing app registration '$APP_NAME' ($app_id)" >&2
  else
    log "Creating app registration '$APP_NAME' ..." >&2
    app_id=$(az ad app create --display-name "$APP_NAME" --sign-in-audience AzureADMyOrg \
      --is-fallback-public-client true --query appId -o tsv) || die "App creation failed"
  fi

  az ad app permission add --id "$app_id" --api "$GRAPH_APP_ID" --api-permissions "${perm_args[@]}" \
    || die "Failed to add Graph API permissions"
  az ad app permission add --id "$app_id" --api "$ARM_APP_ID" --api-permissions "$arm_scope_id=Scope" \
    || die "Failed to add ARM API permission"

  local consent_ok=false i
  for i in 1 2 3 4 5; do
    if az ad app permission admin-consent --id "$app_id" 2>/tmp/.b2c_consent_err.$$; then
      consent_ok=true; rm -f /tmp/.b2c_consent_err.$$; break
    fi
    sleep 10
  done
  if ! $consent_ok; then
    cat /tmp/.b2c_consent_err.$$ >&2 2>/dev/null; rm -f /tmp/.b2c_consent_err.$$
    die "admin-consent failed after retries. Run manually: az ad app permission admin-consent --id $app_id"
  fi

  az ad app update --id "$app_id" --is-fallback-public-client true >/dev/null

  echo "{\"clientId\": \"$app_id\", \"setupAt\": \"$(date -u +%FT%TZ)\"}" > "$marker"
  log "Automation app ready: $app_id" >&2
  echo "$app_id"
}

# ============================================================================
# Graph REST helpers (token-based, work for either the automation app or az)
# ============================================================================
GRAPH_TOKEN=""

_graph_request() {
  local method="$1" url="$2"
  if [[ -n "$GRAPH_TOKEN" ]]; then
    local raw http_code body
    raw=$(curl -sS -w '\n__HTTP_CODE__%{http_code}' -X "${method^^}" "$url" \
      -H "Authorization: Bearer $GRAPH_TOKEN" -H "Content-Type: application/json")
    http_code=$(grep -o '__HTTP_CODE__[0-9]*$' <<<"$raw" | grep -o '[0-9]*$')
    body=$(sed 's/__HTTP_CODE__[0-9]*$//' <<<"$raw")
    if [[ -z "$http_code" || "$http_code" -ge 400 ]]; then echo "$body" >&2; return 1; fi
    echo "$body"; return 0
  else
    local out err_body rc
    err_body=$(mktemp)
    out=$(az rest --method "$method" --url "$url" -o json 2>"$err_body")
    rc=$?
    if [[ $rc -ne 0 ]]; then cat "$err_body" >&2; rm -f "$err_body"; return 1; fi
    rm -f "$err_body"; echo "$out"; return 0
  fi
}

_note_if_permission_error() {
  local err_text="$1"
  if grep -qiE 'Authorization_RequestDenied|Insufficient privileges|Forbidden|does not have any of the required delegated permissions' <<<"$err_text"; then
    err "  => Missing Graph API permission/consent, not a 'nothing to delete' result."
    err "     Try --force-setup to redo the automation app's permission grants."
  fi
}

graph_get_all() {
  local url="$1" out="[]"
  while [[ -n "$url" && "$url" != "null" ]]; do
    local resp err_text
    if ! resp=$(_graph_request get "$url" 2>/tmp/.b2c_err.$$); then
      err_text=$(cat /tmp/.b2c_err.$$ 2>/dev/null); rm -f /tmp/.b2c_err.$$
      err "Graph GET failed: $url"; err "  $err_text"; _note_if_permission_error "$err_text"
      return 1
    fi
    rm -f /tmp/.b2c_err.$$
    out=$(jq -c --argjson a "$out" --argjson b "$(echo "$resp" | jq '.value // []')" -n '$a + $b')
    url=$(echo "$resp" | jq -r '."@odata.nextLink" // empty')
  done
  echo "$out"
}
graph_get_all_beta() { graph_get_all "$GRAPH_API/beta/$1"; }

graph_delete_each() {
  local base="$1" ids_json="$2" label="$3" ver="${4:-v1.0}"
  local count; count=$(echo "$ids_json" | jq 'length')
  if [[ "$count" -eq 0 ]]; then log "No $label found."; return 0; fi
  log "Found $count $label."
  local i=0 failures=0 id name
  while [[ $i -lt $count ]]; do
    id=$(echo "$ids_json" | jq -r ".[$i].id")
    name=$(echo "$ids_json" | jq -r ".[$i].displayName // .[$i].userPrincipalName // .[$i].id")
    if $DRY_RUN; then
      log "  [dry-run] would delete $label: $name ($id)"
    else
      log "  deleting $label: $name ($id)"
      if ! _graph_request delete "$GRAPH_API/$ver/$base/$id" >/dev/null 2>/tmp/.b2c_del_err.$$; then
        err "  failed to delete $label $id: $(cat /tmp/.b2c_del_err.$$ 2>/dev/null)"; rm -f /tmp/.b2c_del_err.$$
        failures=$((failures+1))
      fi
    fi
    i=$((i+1))
  done
  [[ $failures -gt 0 ]] && { err "$failures $label deletion(s) failed."; return 1; }
  return 0
}
graph_delete_each_beta() { graph_delete_each "$1" "$2" "$3" "beta"; }

strip_directory_roles() {
  local user_id="$1"
  local roles; roles=$(_graph_request get "$GRAPH_API/v1.0/users/$user_id/memberOf?\$select=id,displayName" 2>/dev/null) || return 0
  local role_ids; role_ids=$(echo "$roles" | jq -c '[.value[]? | select(."@odata.type" == "#microsoft.graph.directoryRole")]')
  local count; count=$(echo "$role_ids" | jq 'length')
  [[ "$count" -eq 0 ]] && return 0
  log "  user $user_id holds $count directory role(s); removing before delete"
  local i=0 rid rname
  while [[ $i -lt $count ]]; do
    rid=$(echo "$role_ids" | jq -r ".[$i].id"); rname=$(echo "$role_ids" | jq -r ".[$i].displayName")
    if $DRY_RUN; then
      log "    [dry-run] would remove role '$rname' from user $user_id"
    else
      log "    removing role '$rname' from user $user_id"
      _graph_request delete "$GRAPH_API/v1.0/directoryRoles/$rid/members/$user_id/\$ref" >/dev/null 2>/tmp/.b2c_role_err.$$ \
        || err "    failed to remove role '$rname': $(cat /tmp/.b2c_role_err.$$ 2>/dev/null)"
      rm -f /tmp/.b2c_role_err.$$
    fi
    i=$((i+1))
  done
}

graph_delete_users() {
  local ids_json="$1"
  local count; count=$(echo "$ids_json" | jq 'length')
  if [[ "$count" -eq 0 ]]; then log "No user found."; return 0; fi
  log "Found $count user."
  local i=0 failures=0 id name
  while [[ $i -lt $count ]]; do
    id=$(echo "$ids_json" | jq -r ".[$i].id")
    name=$(echo "$ids_json" | jq -r ".[$i].userPrincipalName // .[$i].id")
    strip_directory_roles "$id"
    if $DRY_RUN; then
      log "  [dry-run] would delete user: $name ($id)"
    else
      log "  deleting user: $name ($id)"
      if ! _graph_request delete "$GRAPH_API/v1.0/users/$id" >/dev/null 2>/tmp/.b2c_del_err.$$; then
        err "  failed to delete user $id: $(cat /tmp/.b2c_del_err.$$ 2>/dev/null)"; rm -f /tmp/.b2c_del_err.$$
        failures=$((failures+1))
      fi
    fi
    i=$((i+1))
  done
  [[ $failures -gt 0 ]] && { err "$failures user deletion(s) failed."; return 1; }
  return 0
}

# ============================================================================
# Per-tenant run: cleanup (steps 1-9) + tenant delete (step 10)
# ============================================================================
process_tenant() {
  local B2C_TENANT="$1" SUBSCRIPTION_ID="$2" RESOURCE_GROUP="$3" MGMT_TENANT="$4"
  local APP_CLIENT_ID="$5" MGMT_APP_CLIENT_ID="$6"
  [[ -z "$MGMT_TENANT" ]] && MGMT_TENANT="$B2C_TENANT"
  LOG_PREFIX=" [$B2C_TENANT]"
  GRAPH_TOKEN=""

  if ! $AUTO_YES && ! $DRY_RUN; then
    echo
    echo "This will PERMANENTLY delete tenant resources and then the tenant itself:"
    echo "  B2C tenant:      $B2C_TENANT"
    echo "  Subscription:    $SUBSCRIPTION_ID"
    echo "  Resource group:  $RESOURCE_GROUP"
    echo
    read -r -p "Type the tenant name/id exactly to confirm: " CONFIRM
    if [[ "$CONFIRM" != "$B2C_TENANT" ]]; then err "Confirmation did not match. Skipping this tenant."; return 1; fi
  fi

  if [[ -z "$APP_CLIENT_ID" ]]; then
    APP_CLIENT_ID=$(ensure_automation_app "$B2C_TENANT") || return 1
  fi

  log "Acquiring Graph token via app $APP_CLIENT_ID (reusing cached sign-in if available) ..."
  GRAPH_TOKEN=$(get_token "$APP_CLIENT_ID" "$B2C_TENANT" \
    "https://graph.microsoft.com/IdentityUserFlow.ReadWrite.All https://graph.microsoft.com/Policy.ReadWrite.TrustFramework https://graph.microsoft.com/IdentityProvider.ReadWrite.All https://graph.microsoft.com/TrustFrameworkKeySet.ReadWrite.All https://graph.microsoft.com/Application.ReadWrite.All https://graph.microsoft.com/User.ReadWrite.All https://graph.microsoft.com/RoleManagement.ReadWrite.Directory")
  [[ -z "$GRAPH_TOKEN" ]] && { err "Failed to acquire Graph token via app $APP_CLIENT_ID"; return 2; }
  log "Token acquired."

  local CLEANUP_FAILED=false CURRENT_USER users apps sps idps flows keys policies
  CURRENT_USER=$(_graph_request get "$GRAPH_API/v1.0/me?\$select=id" 2>/dev/null | jq -r '.id // empty')

  log "Step 1/9: Users"
  if users=$(graph_get_all "$GRAPH_API/v1.0/users?\$select=id,userPrincipalName&\$top=999"); then
    users=$(echo "$users" | jq --arg me "$CURRENT_USER" '[.[] | select(.id != $me)]')
    graph_delete_users "$users" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  log "Step 2/9: App registrations (including b2c-extensions-app / IEF apps)"
  local AUTOMATION_APP_OBJECT_ID=""
  if apps=$(graph_get_all "$GRAPH_API/v1.0/applications?\$select=id,appId,displayName&\$top=999"); then
    AUTOMATION_APP_OBJECT_ID=$(echo "$apps" | jq -r --arg cid "$APP_CLIENT_ID" '[.[] | select(.appId == $cid)][0].id // empty')
    apps=$(echo "$apps" | jq --arg cid "$APP_CLIENT_ID" '[.[] | select(.appId != $cid)]')
    graph_delete_each "applications" "$apps" "app registration" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  log "Step 3/9: Enterprise applications (service principals)"
  if sps=$(graph_get_all "$GRAPH_API/v1.0/servicePrincipals?\$select=id,appId,displayName&\$top=999"); then
    sps=$(echo "$sps" | jq --arg cid "$APP_CLIENT_ID" '[.[] | select(.appId != $cid)]')
    if [[ -n "$SP_NAME_FILTER" ]]; then
      local total_before kept
      total_before=$(echo "$sps" | jq 'length')
      sps=$(echo "$sps" | jq --arg names "$SP_NAME_FILTER" \
        '($names | ascii_downcase | split(",") | map(ltrimstr(" ") | rtrimstr(" "))) as $wanted
         | [.[] | select((.displayName // "" | ascii_downcase) as $n | $wanted | index($n) != null)]')
      kept=$(echo "$sps" | jq 'length')
      log "  --sp-name-filter '$SP_NAME_FILTER' applied: $kept of $total_before match, only those will be touched"
      [[ "$kept" -lt "$total_before" ]] && log "  WARNING: $((total_before - kept)) other enterprise application(s) left behind - tenant delete will likely still fail until you widen/drop the filter."
    fi
    graph_delete_each "servicePrincipals" "$sps" "enterprise application" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  log "Step 4/9: Identity providers"
  if idps=$(graph_get_all "$GRAPH_API/v1.0/identity/identityProviders"); then
    graph_delete_each "identity/identityProviders" "$idps" "identity provider" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  log "Step 5/9: User flows"
  if flows=$(graph_get_all_beta "identity/b2cUserFlows"); then
    graph_delete_each_beta "identity/b2cUserFlows" "$flows" "user flow" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  log "Step 6/9: Identity Experience Framework policy keys"
  if keys=$(graph_get_all_beta "trustFramework/keySets"); then
    graph_delete_each_beta "trustFramework/keySets" "$keys" "IEF policy key" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  log "Step 7/9: Identity Experience Framework custom policies"
  if policies=$(graph_get_all_beta "trustFramework/policies"); then
    graph_delete_each_beta "trustFramework/policies" "$policies" "IEF custom policy" || CLEANUP_FAILED=true
  else
    CLEANUP_FAILED=true
  fi

  if [[ -n "$AUTOMATION_APP_OBJECT_ID" ]]; then
    log "Step 8/9: Removing automation app (no longer needed)"
    if $DRY_RUN; then
      log "  [dry-run] would delete app registration: $APP_CLIENT_ID ($AUTOMATION_APP_OBJECT_ID)"
    else
      if ! _graph_request delete "$GRAPH_API/v1.0/applications/$AUTOMATION_APP_OBJECT_ID" >/dev/null 2>/tmp/.b2c_selfdel_err.$$; then
        err "  failed to delete automation app $AUTOMATION_APP_OBJECT_ID: $(cat /tmp/.b2c_selfdel_err.$$ 2>/dev/null)"
        rm -f /tmp/.b2c_selfdel_err.$$; CLEANUP_FAILED=true
      else
        rm -f /tmp/.b2c_selfdel_err.$$
        log "  deleted automation app $AUTOMATION_APP_OBJECT_ID (its service principal is deleted automatically with it)"
        rm -f "$(marker_file_for_tenant "$B2C_TENANT")" \
              "$CACHE_DIR/tokencache-${B2C_TENANT}-${APP_CLIENT_ID}.bin" 2>/dev/null
      fi
    fi
  fi

  if $CLEANUP_FAILED; then err "One or more cleanup steps failed. Not attempting tenant deletion."; return 2; fi
  if $DRY_RUN; then log "Dry run complete. No resources were deleted, tenant was not touched."; return 0; fi

  # -------- Step 9/9: delete the tenant itself (ARM) --------
  local ARM_URL="$ARM_API/subscriptions/$SUBSCRIPTION_ID/resourceGroups/$RESOURCE_GROUP/providers/Microsoft.AzureActiveDirectory/b2cDirectories/$B2C_TENANT?api-version=2021-04-01"
  local ARM_AUTH_APP="" b2c_guid mgmt_guid
  b2c_guid=$(resolve_tenant_guid "$B2C_TENANT")
  mgmt_guid=$(resolve_tenant_guid "$MGMT_TENANT")
  if [[ "$mgmt_guid" == "$b2c_guid" ]]; then
    ARM_AUTH_APP="$APP_CLIENT_ID"
  elif [[ -n "$MGMT_APP_CLIENT_ID" ]]; then
    ARM_AUTH_APP="$MGMT_APP_CLIENT_ID"
  else
    ARM_AUTH_APP=$(ensure_automation_app "$MGMT_TENANT") || return 3
  fi

  log "Acquiring ARM token via app $ARM_AUTH_APP for tenant $MGMT_TENANT (reusing cached sign-in if available) ..."
  local ARM_TOKEN
  ARM_TOKEN=$(get_token "$ARM_AUTH_APP" "$MGMT_TENANT" "https://management.azure.com/user_impersonation")
  [[ -z "$ARM_TOKEN" ]] && { err "Failed to acquire ARM token via app $ARM_AUTH_APP"; return 3; }

  log "Requesting tenant deletion: $ARM_URL"
  local RAW HTTP_CODE BODY
  RAW=$(curl -sS -w '\n__HTTP_CODE__%{http_code}' -X DELETE "$ARM_URL" \
    -H "Authorization: Bearer $ARM_TOKEN" -H "Content-Type: application/json")
  HTTP_CODE=$(grep -o '__HTTP_CODE__[0-9]*$' <<<"$RAW" | grep -o '[0-9]*$')
  BODY=$(sed 's/__HTTP_CODE__[0-9]*$//' <<<"$RAW")
  if [[ -z "$HTTP_CODE" || "$HTTP_CODE" -ge 400 ]]; then
    err "ARM delete call failed ($HTTP_CODE): $BODY"
    if grep -qi "AuthorizationFailed" <<<"$BODY"; then
      err "  => This is an Azure RBAC gap, not a B2C/Graph permission problem. Grant yourself Contributor or Owner on the resource group:"
      err "     az role assignment create --assignee <your-upn> --role Contributor --resource-group $RESOURCE_GROUP"
    fi
    return 3
  fi

  log "Tenant deletion accepted. Polling for completion ..."
  local i CHECK_CODE
  for i in $(seq 1 30); do
    sleep 20
    CHECK_CODE=$(curl -sS -o /dev/null -w '%{http_code}' "$ARM_URL" -H "Authorization: Bearer $ARM_TOKEN")
    if [[ "$CHECK_CODE" == "404" ]]; then log "Tenant resource no longer exists. Deletion complete."; return 0; fi
    log "  still deleting... (check $i/30, last status $CHECK_CODE)"
  done
  err "Timed out waiting for tenant deletion to finish. Check the Azure portal for final status."
  return 3
}

# ============================================================================
# Main
# ============================================================================
preflight
write_token_helper

if [[ -n "$CSV_FILE" ]]; then
  [[ -f "$CSV_FILE" ]] || die "CSV file not found: $CSV_FILE"
  RESULTS="results-$(date '+%Y%m%d-%H%M%S').log"
  echo "tenant,status,detail" > "$RESULTS"
  while IFS=',' read -r c_tenant c_sub c_rg c_mgmt c_app c_mgmt_app; do
    [[ -z "$c_tenant" || "$c_tenant" =~ ^# ]] && continue
    c_tenant=$(echo "$c_tenant" | xargs); c_sub=$(echo "$c_sub" | xargs); c_rg=$(echo "$c_rg" | xargs)
    c_mgmt=$(echo "${c_mgmt:-}" | xargs); c_app=$(echo "${c_app:-}" | xargs); c_mgmt_app=$(echo "${c_mgmt_app:-}" | xargs)
    echo; echo "==================================================================="; echo "Processing tenant: $c_tenant"; echo "==================================================================="
    if process_tenant "$c_tenant" "$c_sub" "$c_rg" "$c_mgmt" "$c_app" "$c_mgmt_app"; then
      echo "$c_tenant,SUCCESS," >> "$RESULTS"
    else
      rc=$?
      echo "$c_tenant,FAILED,exit_code_$rc" >> "$RESULTS"
      echo "!! Failed to process $c_tenant (exit $rc) - continuing with next tenant"
    fi
  done < "$CSV_FILE"
  echo; echo "Done. Results written to $RESULTS"
  column -s, -t "$RESULTS" 2>/dev/null || cat "$RESULTS"
else
  process_tenant "$B2C_TENANT" "$SUBSCRIPTION_ID" "$RESOURCE_GROUP" "$MGMT_TENANT" "$APP_CLIENT_ID" "$MGMT_APP_CLIENT_ID"
  exit $?
fi

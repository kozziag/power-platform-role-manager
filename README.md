# Power Platform Role Manager (hosted, multi-tenant)

A client-side-only web app for assigning Power Platform RBAC role
assignments (Power Platform Reader, Contributor, Owner, RBAC Administrator,
etc.) to Entra ID users or applications. **No backend server.** Everyone who
opens the link signs in with their own organization's Entra ID; tokens live
only in their own browser's `sessionStorage` and are never seen by any
server this project controls.

Live app: **https://kozziag.github.io/power-platform-role-manager/**

Prefer PowerShell instead of a web browser? See
[`scripts/Invoke-PowerPlatformRoleManager.ps1`](scripts/Invoke-PowerPlatformRoleManager.ps1)
-- an interactive console version with the same scoped lookups and live
activity feedback, requiring no custom app registration or admin consent at
all (see its own section below).

## Why this architecture

Both Microsoft Graph and the Power Platform Authorization API were verified
live to send `Access-Control-Allow-Origin: *` on their CORS preflight
responses, so a static page can call them directly from the browser with no
backend at all. That's a strictly better security posture than a hosted
backend would be for a multi-tenant admin tool: there's no shared server
state for one user's session to ever leak into another's.

## App registration

This app is registered in Entra ID as:

- **Name:** Power Platform Role Manager
- **App (client) ID:** `2fad7640-e268-49b8-8a9c-0ea1b0bf0921`
- **Type:** Single-page application (public client, Authorization Code + PKCE
  — no client secret, nothing sensitive to leak from this repo)
- **Sign-in audience:** `AzureADMultipleOrgs` (any Entra tenant)
- **Redirect URIs:** the GitHub Pages URL above, plus `http://localhost:5500/`
  and `http://localhost:8080/` for local testing
- **API permissions requested (delegated, least-privilege — not a blanket
  `.default`):**
  - Microsoft Graph: `User.Read.All` (to search/resolve people and service
    principals)
  - Power Platform API (`https://api.powerplatform.com`):
    `Authorization.RoleAssignments.Read`, `Authorization.RoleAssignments.Write`

## First-time setup for a new organization (one-time admin consent)

Because `User.Read.All` and the Power Platform role-assignment scopes are
admin-restricted, **a Global Administrator, Privileged Role Administrator, or
Application Administrator in your tenant must consent once** before anyone
in your org can use this tool. Have that admin visit:

```
https://login.microsoftonline.com/organizations/adminconsent?client_id=2fad7640-e268-49b8-8a9c-0ea1b0bf0921&redirect_uri=https://kozziag.github.io/power-platform-role-manager/
```

After that one-time consent, anyone in that tenant can sign in and use the
tool immediately — no further setup.

## What the app can and cannot do

Signing in does **not** grant you any new permission. Whether an assignment
actually succeeds is enforced entirely server-side by the real Power
Platform Authorization API, based on the signed-in user's own existing RBAC
permissions in their tenant (typically Power Platform RBAC Administrator or
Global Administrator). This app is a UI over that API, not a privilege
escalation path.

## Scoped lookups by design

The "current roles" list is deliberately scoped to **one principal at a
time** (`$filter=principalObjectId eq <guid>`, confirmed to work server-side
against the live API — note the GUID must be unquoted; a quoted literal
returns 400). It does **not** support listing every role assignment in the
tenant, because that does not scale in orgs with large numbers of
assignments. A client-side "recent activity" log shows what changed in your
own session in real time instead, with no extra network calls.

## PowerShell alternative (no app registration, no admin consent)

`scripts/Invoke-PowerPlatformRoleManager.ps1` is an interactive console
version of this same tool for admins who prefer PowerShell over a browser.

**Why it needs no setup at all (unlike the web app above):** it signs you in
via `Connect-AzAccount`, which uses the Az PowerShell module's own
pre-consented, first-party multi-tenant app registration to get tokens for
Microsoft Graph and the Power Platform API -- the same mechanism the Azure
CLI uses under the hood. There's no custom app registration to register and
no tenant admin consent step to complete before you can run it.

Features:
- Interactive menu-driven flow with colorized, real-time feedback on every
  action (`[OK]` / `[!!]` / `[ERR]` prefixes, spinners during API calls)
- Search Entra users, or resolve a service principal from an App ID
- View a principal's existing roles (same scoped `$filter`, not a tenant-wide
  dump) before assigning
- Explicit `type 'yes' to confirm` prompt before every assign/remove --
  nothing destructive happens without an explicit, typed confirmation
- A running, in-session activity log (`Format-Table`) of everything you changed
- Real, unmodified API error messages surfaced on failure -- never a guess

Requirements: PowerShell 5.1+ (or 7+), the `Az.Accounts` module (the script
offers to install it for you), and sufficient Power Platform RBAC permission
in your own tenant to actually assign/view roles -- same permission model as
the web app.

```powershell
.\scripts\Invoke-PowerPlatformRoleManager.ps1
```

## Local development

Just serve the `web/` folder with any static file server, e.g.:

```
cd web
python -m http.server 8080
```

Then open `http://localhost:8080/` (already registered as a valid redirect
URI above).

## Files

- `web/index.html` — markup + styles
- `web/app.js` — all logic: MSAL.js auth, Graph/Power Platform API calls,
  modal dialogs (no native browser popups), scoped assignment lookups,
  real-time activity log
- `scripts/Invoke-PowerPlatformRoleManager.ps1` — PowerShell console
  alternative, no app registration required

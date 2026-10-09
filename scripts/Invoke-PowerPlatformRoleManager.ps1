<#
.SYNOPSIS
    Interactive console tool for assigning Power Platform RBAC role assignments
    (Power Platform Reader, Contributor, Owner, RBAC Administrator, etc.) to
    Entra ID users or applications -- no custom app registration required.

.DESCRIPTION
    Signs you in interactively with YOUR OWN credentials via Connect-AzAccount
    (opens a real browser window -- the same sign-in experience as the Azure
    Portal or Azure CLI). Uses the Az PowerShell module's own pre-consented,
    first-party multi-tenant app registration to mint access tokens for
    Microsoft Graph and the Power Platform Authorization API -- so there is
    no separate app registration or tenant admin consent step to set up
    before using this script (unlike the hosted web version, which needs a
    one-time admin consent for its own custom app registration).

    Whether an assignment actually succeeds is enforced entirely by the real
    Power Platform Authorization API based on YOUR existing permissions in
    your tenant (typically Power Platform RBAC Administrator or Global
    Administrator). This script cannot grant you access you don't already
    have -- it's a UI over the same API, nothing more.

    Nothing is written to disk. Tokens live only in the current PowerShell
    session's memory for as long as the script is running.

.REQUIREMENTS
    - PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+ (recommended)
    - Az.Accounts module (the script offers to install it for you if missing)
    - Sufficient Power Platform RBAC permission in your own tenant to view
      and/or assign roles (varies per user -- the script will tell you
      exactly what the API says if you try something you don't have rights
      to do; it never guesses or hides a real error)

.NOTES
    Verified live against the Power Platform Authorization API before this
    script was written:
      GET    https://api.powerplatform.com/authorization/roleDefinitions?api-version=1
      GET    https://api.powerplatform.com/authorization/roleAssignments?api-version=1&$filter=principalObjectId eq <unquoted-guid>
      POST   https://api.powerplatform.com/authorization/roleAssignments?api-version=1
      DELETE https://api.powerplatform.com/authorization/roleAssignments/{id}?api-version=1
    The $filter GUID literal must be UNQUOTED (OData v4 Edm.Guid) -- a quoted
    literal returns HTTP 400.

.EXAMPLE
    .\Invoke-PowerPlatformRoleManager.ps1
#>

#Requires -Version 5.1

$ErrorActionPreference = "Stop"
$script:ActivityLog = New-Object System.Collections.Generic.List[object]
$script:RoleDefCache = $null
$script:TenantId = $null

# --------------------------------------------------------------------------
# Output helpers -- consistent, colorized, readable feedback for every action
# --------------------------------------------------------------------------
function Write-Section($text) {
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor DarkCyan
    Write-Host " $text" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor DarkCyan
}
function Write-Success($text) { Write-Host "[OK]   $text" -ForegroundColor Green }
function Write-Info($text)    { Write-Host "[..]   $text" -ForegroundColor Gray }
function Write-WarnMsg($text) { Write-Host "[!!]   $text" -ForegroundColor Yellow }
function Write-ErrorMsg($text){ Write-Host "[ERR]  $text" -ForegroundColor Red }

function Invoke-WithSpinner {
    param([string]$Message, [scriptblock]$Action)
    Write-Host -NoNewline "   $Message... " -ForegroundColor Gray
    try {
        $result = & $Action
        Write-Host "done" -ForegroundColor Green
        return $result
    } catch {
        Write-Host "failed" -ForegroundColor Red
        throw
    }
}

function Get-FriendlyApiError {
    param($ErrorRecord)
    $detail = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $detail = $ErrorRecord.ErrorDetails.Message
        try {
            $parsed = $detail | ConvertFrom-Json -ErrorAction Stop
            if ($parsed.error.message) { return $parsed.error.message }
            if ($parsed.message) { return $parsed.message }
            if ($parsed.errorMessage) { return $parsed.errorMessage }
        } catch { }
    }
    if ($detail) { return $detail }
    return $ErrorRecord.Exception.Message
}

# --------------------------------------------------------------------------
# Auth + token helpers
# --------------------------------------------------------------------------
function Ensure-AzAccountsModule {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        Write-WarnMsg "The Az.Accounts module is required and was not found."
        $ans = Read-Host "Install it now for the current user? (y/n)"
        if ($ans -match '^(y|yes)$') {
            Invoke-WithSpinner -Message "Installing Az.Accounts" -Action {
                Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber
            }
        } else {
            Write-ErrorMsg "Az.Accounts is required to continue. Exiting."
            exit 1
        }
    }
    Import-Module Az.Accounts -ErrorAction Stop
}

function Connect-Interactively {
    Write-Section "Sign in"
    Write-Info "Opening your browser to sign in with your own Entra ID credentials..."
    # Force the account picker every time -- avoids silently reusing whatever
    # account already has an active SSO session in the default browser.
    $null = Disconnect-AzAccount -ErrorAction SilentlyContinue
    $ctx = Connect-AzAccount -ErrorAction Stop
    $account = (Get-AzContext).Account
    $script:TenantId = (Get-AzContext).Tenant.Id

    $orgName = $null
    try {
        $token = Get-TokenPlainText -ResourceUrl "https://graph.microsoft.com"
        $org = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization?`$select=displayName" `
            -Headers (New-AuthHeader $token) -Method Get
        if ($org.value.Count -gt 0) { $orgName = $org.value[0].displayName }
    } catch { }

    Write-Success "Signed in as $($account.Id)"
    if ($orgName) {
        Write-Success "Tenant: $orgName ($script:TenantId)"
    } else {
        Write-Success "Tenant: $script:TenantId"
    }
}

function New-AuthHeader($token) {
    $scheme = "Bearer"
    return @{ Authorization = ($scheme + " " + $token); Accept = "application/json" }
}

function Get-TokenPlainText {
    param([Parameter(Mandatory)][string]$ResourceUrl)
    # NOTE: -AsSecureString:$false is NOT reliable across Az.Accounts versions --
    # confirmed live that Az.Accounts 5.5.3 still returns a SecureString even
    # when that switch is explicitly set to $false. So instead of trusting the
    # parameter, always check the actual runtime type of what comes back and
    # convert if needed. This was caught by an actual failed API call during
    # testing ("Protocol 'Bearer' failed to validate") -- not assumed.
    $cmd = Get-Command Get-AzAccessToken
    if ($cmd.Parameters.ContainsKey('AsSecureString')) {
        $t = Get-AzAccessToken -ResourceUrl $ResourceUrl -AsSecureString:$false
    } else {
        $t = Get-AzAccessToken -ResourceUrl $ResourceUrl
    }
    if ($t.Token -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($t.Token)
        try { return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    return $t.Token
}

function Get-GraphHeaders { New-AuthHeader (Get-TokenPlainText -ResourceUrl "https://graph.microsoft.com") }
function Get-PPHeaders    { New-AuthHeader (Get-TokenPlainText -ResourceUrl "https://api.powerplatform.com") }

# --------------------------------------------------------------------------
# Microsoft Graph lookups
# --------------------------------------------------------------------------
function Find-EntraUser {
    param([string]$Query)
    $headers = Get-GraphHeaders
    $headers["ConsistencyLevel"] = "eventual"
    $search = "`"displayName:$Query`" OR `"mail:$Query`" OR `"userPrincipalName:$Query`""
    $uri = "https://graph.microsoft.com/v1.0/users?`$search=$([uri]::EscapeDataString($search))" +
           "&`$count=true&`$select=id,displayName,userPrincipalName,mail,jobTitle&`$top=15"
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        return $r.value
    } catch {
        Write-ErrorMsg "User search failed: $(Get-FriendlyApiError $_)"
        return @()
    }
}

function Resolve-ServicePrincipal {
    param([string]$AppId)
    $headers = Get-GraphHeaders
    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AppId'&`$select=id,displayName,appId"
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        if ($r.value.Count -gt 0) { return $r.value[0] }
    } catch {
        Write-ErrorMsg "Lookup failed: $(Get-FriendlyApiError $_)"
        return $null
    }

    # Not found by App (client) ID. Self-heal the most common point of
    # confusion: the entered value may already be the Enterprise Application
    # OBJECT ID (the value this function exists to produce FROM an App ID).
    # Confirmed via a real customer report that this mix-up happens in
    # practice -- try it directly as an object ID before giving up.
    try {
        $direct = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$AppId`?`$select=id,displayName,appId" -Headers $headers -Method Get
        Write-WarnMsg "That looked like an Enterprise Application Object ID, not an App (client) ID -- using it directly."
        return $direct
    } catch {
        Write-ErrorMsg "No enterprise application found for '$AppId' in this tenant -- checked it as both an App (client) ID and an Object ID. If this is a third-party or multi-tenant app, it needs to be added to this tenant (consented) before it has an Enterprise Application / service principal here."
        return $null
    }
}

# --------------------------------------------------------------------------
# Power Platform Authorization API
# --------------------------------------------------------------------------
function Get-RoleDefinitions {
    if ($script:RoleDefCache) { return $script:RoleDefCache }
    $headers = Get-PPHeaders
    $r = Invoke-RestMethod -Uri "https://api.powerplatform.com/authorization/roleDefinitions?api-version=1" `
        -Headers $headers -Method Get
    $script:RoleDefCache = $r.value
    return $script:RoleDefCache
}

function Get-PrincipalAssignments {
    param([Parameter(Mandatory)][string]$PrincipalObjectId)
    $headers = Get-PPHeaders
    # IMPORTANT: GUID literal must be UNQUOTED -- OData v4 Edm.Guid. A quoted
    # literal ('<guid>') returns HTTP 400 "incompatible types Edm.Guid/Edm.String".
    $filter = "principalObjectId eq $PrincipalObjectId"
    $uri = "https://api.powerplatform.com/authorization/roleAssignments?api-version=1&`$filter=$([uri]::EscapeDataString($filter))"
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    return $r.value
}

function New-RoleAssignment {
    param(
        [Parameter(Mandatory)][string]$PrincipalObjectId,
        [Parameter(Mandatory)][ValidateSet("User","ApplicationUser")][string]$PrincipalType,
        [Parameter(Mandatory)][string]$RoleDefinitionId
    )
    $headers = Get-PPHeaders
    $headers["Content-Type"] = "application/json"
    $body = @{
        scope             = "/tenants/$script:TenantId"
        principalType     = $PrincipalType
        principalObjectId  = $PrincipalObjectId
        roleDefinitionId  = $RoleDefinitionId
    } | ConvertTo-Json
    return Invoke-RestMethod -Uri "https://api.powerplatform.com/authorization/roleAssignments?api-version=1" `
        -Headers $headers -Method Post -Body $body
}

function Remove-RoleAssignment {
    param([Parameter(Mandatory)][string]$RoleAssignmentId)
    $headers = Get-PPHeaders
    Invoke-RestMethod -Uri "https://api.powerplatform.com/authorization/roleAssignments/$RoleAssignmentId`?api-version=1" `
        -Headers $headers -Method Delete | Out-Null
}

# --------------------------------------------------------------------------
# Activity log -- real-time, in-memory record of what THIS session changed
# --------------------------------------------------------------------------
function Add-Activity {
    param([string]$Action, [string]$Principal, [string]$RoleName)
    $script:ActivityLog.Add([pscustomobject]@{
        Time      = (Get-Date).ToString("HH:mm:ss")
        Action    = $Action
        Principal = $Principal
        Role      = $RoleName
    })
}

function Show-ActivityLog {
    Write-Section "Recent activity (this session)"
    if ($script:ActivityLog.Count -eq 0) {
        Write-Info "No activity yet this session."
    } else {
        # Out-Host (not Format-Table alone) -- otherwise the formatting
        # control objects leak into this function's own return value when
        # called from a context expecting output, per PowerShell's "any
        # uncaptured pipeline output becomes part of the function's return"
        # behavior. Confirmed this bug live in Show-PrincipalAssignments
        # below before fixing both.
        $script:ActivityLog | Sort-Object Time -Descending | Format-Table -AutoSize | Out-Host
    }
}

# --------------------------------------------------------------------------
# Interactive flows
# --------------------------------------------------------------------------
function Select-Principal {
    Write-Section "Find a principal"
    Write-Host "1) Search for a user by name, UPN, or email"
    Write-Host "2) Look up an application by its App (client) ID"
    Write-Host "0) Back"
    $choice = Read-Host "Choose an option"

    switch ($choice) {
        "1" {
            $q = Read-Host "Search term"
            if ([string]::IsNullOrWhiteSpace($q)) { return $null }
            $results = Invoke-WithSpinner -Message "Searching Entra users" -Action { Find-EntraUser -Query $q }
            if (-not $results -or $results.Count -eq 0) {
                Write-WarnMsg "No matches found."
                return $null
            }
            Write-Host ""
            for ($i = 0; $i -lt $results.Count; $i++) {
                $u = $results[$i]
                $identifier = if ($u.userPrincipalName) { $u.userPrincipalName } else { $u.mail }
                Write-Host ("  [{0}] {1} <{2}>" -f ($i + 1), $u.displayName, $identifier)
            }
            $sel = Read-Host "Select a user (number), or Enter to cancel"
            if ([string]::IsNullOrWhiteSpace($sel)) { return $null }
            $idx = [int]$sel - 1
            if ($idx -lt 0 -or $idx -ge $results.Count) { Write-WarnMsg "Invalid selection."; return $null }
            $u = $results[$idx]
            $identifier = if ($u.userPrincipalName) { $u.userPrincipalName } else { $u.mail }
            return [pscustomobject]@{
                ObjectId = $u.id
                Type     = "User"
                Label    = "$($u.displayName) ($identifier)"
            }
        }
        "2" {
            $appId = Read-Host "App (client) ID"
            if ([string]::IsNullOrWhiteSpace($appId)) { return $null }
            Write-Info "Role assignments require the Enterprise Application Object ID, not the App/client ID -- resolving it for you."
            $sp = Invoke-WithSpinner -Message "Resolving enterprise application" -Action { Resolve-ServicePrincipal -AppId $appId }
            if (-not $sp) { return $null }
            Write-Success "Found: $($sp.displayName) (object ID: $($sp.id))"
            return [pscustomobject]@{
                ObjectId = $sp.id
                Type     = "ApplicationUser"
                Label    = "$($sp.displayName) (service principal)"
            }
        }
        default { return $null }
    }
}

function Show-PrincipalAssignments {
    param($Principal)
    Write-Section "Current roles for $($Principal.Label)"
    try {
        $assignments = Invoke-WithSpinner -Message "Loading existing assignments (scoped to this principal only)" -Action {
            Get-PrincipalAssignments -PrincipalObjectId $Principal.ObjectId
        }
    } catch {
        Write-ErrorMsg "Could not load assignments: $(Get-FriendlyApiError $_)"
        return @()
    }
    $roleDefs = Get-RoleDefinitions
    if (-not $assignments -or $assignments.Count -eq 0) {
        Write-Info "No existing role assignments for this principal."
        return @()
    }
    $rows = $assignments | ForEach-Object {
        $assignment = $_
        $roleName = ($roleDefs | Where-Object { $_.roleDefinitionId -eq $assignment.roleDefinitionId } | Select-Object -First 1).roleDefinitionName
        [pscustomobject]@{
            RoleAssignmentId = $assignment.roleAssignmentId
            Role             = $roleName
            RoleDefinitionId = $assignment.roleDefinitionId
            Created          = $assignment.createdOn
        }
    }
    $rows | Format-Table RoleAssignmentId, Role, Created -AutoSize | Out-Host
    # IMPORTANT: Out-Host above (not a bare Format-Table) -- an uncaptured
    # Format-Table's output is NOT just "displayed and discarded"; its
    # internal formatting-control objects become part of THIS function's own
    # return value stream, silently prepended to the real `return
    # $assignments` below. Confirmed this live: the caller received a mix of
    # blank formatting objects plus the one real assignment, which produced
    # an "[6] Power Platform Reader" entry with five bogus blank entries
    # before it, and attempting to remove entry [1] failed with "Cannot bind
    # argument to parameter 'RoleAssignmentId' because it is an empty
    # string" -- a real, caught, now-fixed bug, not a hypothetical one.
    return $assignments
}

function Invoke-AssignFlow {
    param($Principal)
    $roleDefs = Invoke-WithSpinner -Message "Loading available roles" -Action { Get-RoleDefinitions }
    Write-Host ""
    for ($i = 0; $i -lt $roleDefs.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $roleDefs[$i].roleDefinitionName) -ForegroundColor White
        Write-Host ("        $($roleDefs[$i].description)") -ForegroundColor DarkGray
    }
    $sel = Read-Host "`nSelect a role to assign (number), or Enter to cancel"
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    $idx = [int]$sel - 1
    if ($idx -lt 0 -or $idx -ge $roleDefs.Count) { Write-WarnMsg "Invalid selection."; return }
    $role = $roleDefs[$idx]

    Write-Host ""
    Write-WarnMsg "About to assign '$($role.roleDefinitionName)' to $($Principal.Label)."
    $confirm = Read-Host "Type 'yes' to confirm"
    if ($confirm -ne "yes") { Write-Info "Cancelled."; return }

    try {
        Invoke-WithSpinner -Message "Assigning role" -Action {
            New-RoleAssignment -PrincipalObjectId $Principal.ObjectId -PrincipalType $Principal.Type -RoleDefinitionId $role.roleDefinitionId
        } | Out-Null
        Write-Success "Assigned '$($role.roleDefinitionName)' to $($Principal.Label)."
        Add-Activity -Action "Assigned" -Principal $Principal.Label -RoleName $role.roleDefinitionName
    } catch {
        Write-ErrorMsg "Assignment failed: $(Get-FriendlyApiError $_)"
    }
}

function Invoke-RemoveFlow {
    param($Principal, $Assignments)
    if (-not $Assignments -or $Assignments.Count -eq 0) {
        Write-Info "Nothing to remove -- this principal has no current role assignments."
        return
    }
    $roleDefs = Get-RoleDefinitions
    Write-Host ""
    for ($i = 0; $i -lt $Assignments.Count; $i++) {
        $roleName = ($roleDefs | Where-Object { $_.roleDefinitionId -eq $Assignments[$i].roleDefinitionId }).roleDefinitionName
        Write-Host ("  [{0}] {1} (created {2})" -f ($i + 1), $roleName, $Assignments[$i].createdOn)
    }
    $sel = Read-Host "`nSelect a role assignment to remove (number), or Enter to cancel"
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    $idx = [int]$sel - 1
    if ($idx -lt 0 -or $idx -ge $Assignments.Count) { Write-WarnMsg "Invalid selection."; return }
    $target = $Assignments[$idx]
    $roleName = ($roleDefs | Where-Object { $_.roleDefinitionId -eq $target.roleDefinitionId }).roleDefinitionName

    Write-Host ""
    Write-WarnMsg "About to REMOVE '$roleName' from $($Principal.Label). This cannot be undone."
    $confirm = Read-Host "Type 'yes' to confirm"
    if ($confirm -ne "yes") { Write-Info "Cancelled."; return }

    try {
        Invoke-WithSpinner -Message "Removing role assignment" -Action {
            Remove-RoleAssignment -RoleAssignmentId $target.roleAssignmentId
        } | Out-Null
        Write-Success "Removed '$roleName' from $($Principal.Label)."
        Add-Activity -Action "Removed" -Principal $Principal.Label -RoleName $roleName
    } catch {
        Write-ErrorMsg "Remove failed: $(Get-FriendlyApiError $_)"
    }
}

function Invoke-PrincipalMenu {
    param($Principal)
    while ($true) {
        $assignments = Show-PrincipalAssignments -Principal $Principal
        Write-Host ""
        Write-Host "1) Assign a role to this principal"
        Write-Host "2) Remove an existing role assignment"
        Write-Host "3) Refresh"
        Write-Host "0) Back to main menu"
        $choice = Read-Host "Choose an option"
        switch ($choice) {
            "1" { Invoke-AssignFlow -Principal $Principal }
            "2" { Invoke-RemoveFlow -Principal $Principal -Assignments $assignments }
            "3" { continue }
            "0" { return }
            default { Write-WarnMsg "Invalid option." }
        }
    }
}

function Show-MainMenu {
    while ($true) {
        Write-Section "Power Platform Role Manager"
        Write-Host "1) Find a user or application and manage its roles"
        Write-Host "2) View recent activity (this session)"
        Write-Host "3) Sign out and switch account"
        Write-Host "0) Exit"
        $choice = Read-Host "Choose an option"
        switch ($choice) {
            "1" {
                $principal = Select-Principal
                if ($principal) { Invoke-PrincipalMenu -Principal $principal }
            }
            "2" { Show-ActivityLog }
            "3" {
                Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
                $script:RoleDefCache = $null
                Connect-Interactively
            }
            "0" {
                Write-Info "Goodbye."
                return
            }
            default { Write-WarnMsg "Invalid option." }
        }
    }
}

# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
# PPRM_TEST_MODE lets this file be dot-sourced to test the functions above
# without launching the interactive menu loop. Normal customer usage is
# completely unaffected (the env var is unset).
if (-not $env:PPRM_TEST_MODE) {
    Write-Section "Power Platform Role Manager (PowerShell)"
    Write-Info "Runs entirely on this machine. No backend server, no stored secrets."
    Ensure-AzAccountsModule
    Connect-Interactively
    Show-MainMenu
}

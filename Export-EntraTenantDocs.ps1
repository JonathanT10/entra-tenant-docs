<#
.SYNOPSIS
    Generate living documentation of an Entra ID tenant - Markdown, JSON, and
    a self-contained HTML report. Read-only.

.DESCRIPTION
    One run produces three views of the same snapshot:

      docs/          numbered Markdown files - human-readable documentation.
                     Deterministic output (stable sort, no timestamps in the
                     section files), so committing the folder to git turns
                     every re-run into a config-drift diff.
      tenant.json    the complete structured snapshot - the data contract
                     for anything you build on top.
      report.html    a self-contained, timestamped report for people who
                     will never open a markdown file - KPI tiles, license
                     meters, credential-expiry status, Conditional Access
                     at a glance. No server, no dependencies; open in any
                     browser, drop on any intranet share.

    What gets documented (identity plane, v1):
      1. Tenant overview - org info, verified domains, license SKUs
      2. Conditional Access - every policy rendered readable + named locations
      3. Directory roles - permanent assignments per role
      4. Groups - counts, role-assignable groups, dynamic membership rules
      5. Authentication methods policy
      6. User & guest settings (authorization policy) in plain English
      7. App registrations - sign-in audience and credential expiry

    Makes no changes to the tenant. Every call is a read.

.PARAMETER OutputPath
    Folder for the output (created if missing). Default: .\tenant-docs

.PARAMETER FromJson
    Re-render docs + report from a previously exported tenant.json instead
    of calling Graph. No connection needed.

.PARAMETER SampleData
    Render from the bundled sample-data.json - see what the output looks
    like with zero permissions and zero tenants harmed.

.PARAMETER StaleCredDays
    Days-until-expiry threshold under which an app credential is flagged
    as expiring soon. Default: 90 (30 = "serious" internally).

.EXAMPLE
    .\Export-EntraTenantDocs.ps1

    Documents the connected tenant into .\tenant-docs\

.EXAMPLE
    .\Export-EntraTenantDocs.ps1 -SampleData -OutputPath .\demo

    Renders the bundled sample tenant - no Graph connection at all.

.NOTES
    Required Graph scopes (all read-only):
        Directory.Read.All, Policy.Read.All,
        RoleManagement.Read.Directory, Application.Read.All
    Required module:
        Microsoft.Graph.Authentication (only)

    PIM eligible role assignments are not included (permanent only).
    Exchange / Intune / SharePoint settings are out of scope for v1.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = '.\tenant-docs',
    [string]$FromJson,
    [switch]$SampleData,
    [ValidateRange(1, 3650)][int]$StaleCredDays = 90
)

$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

function Get-GraphPage {
    param([string]$Uri, [hashtable]$Headers)
    $rows = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = if ($Headers) {
            Invoke-MgGraphRequest -Method GET -Uri $next -Headers $Headers -OutputType PSObject
        } else {
            Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        }
        if ($resp.value) { foreach ($r in $resp.value) { $rows.Add($r) } }
        $next = $resp.'@odata.nextLink'
    }
    return $rows
}

function Get-GraphCount {
    # @odata.count via $count=true + ConsistencyLevel: eventual
    param([string]$Resource, [string]$Filter)
    $uri = "https://graph.microsoft.com/v1.0/$Resource`?`$count=true&`$top=1"
    if ($Filter) { $uri += "&`$filter=" + [uri]::EscapeDataString($Filter) }
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri `
        -Headers @{ ConsistencyLevel = 'eventual' } -OutputType PSObject
    return [int]$resp.'@odata.count'
}

function MdEscape { param($s) if ($null -eq $s) { return '' } ($s -replace '\|', '\|') -replace "`r`n|`n", ' ' }

function Iso {
    # Graph datetimes arrive as [datetime] (PSObject conversion) or ISO strings;
    # normalize to a round-trippable UTC ISO string, culture-independent.
    param($v)
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($v -is [datetime]) { return $v.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    try { return ([datetime]::Parse("$v", [cultureinfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AdjustToUniversal)).ToString('yyyy-MM-ddTHH:mm:ssZ') }
    catch { return "$v" }
}

# --------------------------------------------------------------------------- #
# Collection (Graph -> $data). Skipped entirely in -FromJson / -SampleData mode.
# --------------------------------------------------------------------------- #

function Get-TenantData {
    param([int]$StaleCredDays)

    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'Directory.Read.All','Policy.Read.All',
            'RoleManagement.Read.Directory','Application.Read.All' -NoWelcome
    }

    $G = 'https://graph.microsoft.com/v1.0'

    # -- 1. Organization + licenses ----------------------------------------- #
    Write-Host 'Collecting organization + licenses...'
    $org = (Get-GraphPage "$G/organization")[0]
    $domains = @($org.verifiedDomains | Sort-Object name | ForEach-Object {
        [ordered]@{ Name = $_.name; IsDefault = [bool]$_.isDefault; IsInitial = [bool]$_.isInitial }
    })
    $skus = Get-GraphPage "$G/subscribedSkus"
    $licenses = @($skus | Sort-Object skuPartNumber | ForEach-Object {
        $purchased = [int]$_.prepaidUnits.enabled
        [ordered]@{
            Sku       = $_.skuPartNumber
            Purchased = $purchased
            Assigned  = [int]$_.consumedUnits
            Available = [Math]::Max(0, $purchased - [int]$_.consumedUnits)
        }
    })

    Write-Host 'Counting users...'
    $userCounts = [ordered]@{
        Members         = Get-GraphCount 'users' "userType eq 'Member'"
        EnabledMembers  = Get-GraphCount 'users' "userType eq 'Member' and accountEnabled eq true"
        Guests          = Get-GraphCount 'users' "userType eq 'Guest'"
    }

    # -- 2. Conditional Access ---------------------------------------------- #
    Write-Host 'Collecting Conditional Access policies...'
    $rawPolicies = Get-GraphPage "$G/identity/conditionalAccess/policies"
    $rawLocations = Get-GraphPage "$G/identity/conditionalAccess/namedLocations"
    $locNameById = @{}
    foreach ($l in $rawLocations) { $locNameById[$l.id] = $l.displayName }

    # Role definitions double as the CA role-id map (CA references role template ids)
    Write-Host 'Collecting directory roles...'
    $roleDefs = Get-GraphPage "$G/roleManagement/directory/roleDefinitions?`$select=id,displayName,templateId"
    $roleNameById = @{}
    foreach ($d in $roleDefs) {
        $roleNameById[$d.id] = $d.displayName
        if ($d.templateId) { $roleNameById[$d.templateId] = $d.displayName }
    }

    # Resolve every user/group GUID referenced by CA policies in one POST
    $refIds = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in $rawPolicies) {
        $u = $p.conditions.users
        foreach ($col in @($u.includeUsers, $u.excludeUsers, $u.includeGroups, $u.excludeGroups)) {
            foreach ($id in @($col)) {
                if ($id -and $id -match '^[0-9a-f-]{36}$') { [void]$refIds.Add($id) }
            }
        }
    }
    $dirNameById = @{}
    if ($refIds.Count) {
        Write-Host "Resolving $($refIds.Count) directory object ids..."
        try {
            $body = @{ ids = @($refIds); types = @('user', 'group') } | ConvertTo-Json
            $resolved = Invoke-MgGraphRequest -Method POST -Uri "$G/directoryObjects/getByIds" `
                -Body $body -ContentType 'application/json' -OutputType PSObject
            foreach ($o in @($resolved.value)) { $dirNameById[$o.id] = $o.displayName }
        }
        catch { Write-Warning "getByIds failed ($($_.Exception.Message)) - GUIDs left unresolved." }
    }

    $WELL_KNOWN_APPS = @{
        'All' = 'All cloud apps'; 'None' = 'None'; 'Office365' = 'Office 365'
        'MicrosoftAdminPortals' = 'Microsoft Admin Portals'
    }
    function Resolve-CaList {
        param($Ids, [hashtable]$Map, [hashtable]$Fallback)
        @(foreach ($id in @($Ids)) {
            if ($null -eq $id -or $id -eq '') { continue }
            if ($Fallback -and $Fallback.ContainsKey($id)) { $Fallback[$id] }
            elseif ($Map -and $Map.ContainsKey($id)) { $Map[$id] }
            elseif ($id -in 'All', 'None', 'GuestsOrExternalUsers') { $id }
            else { $id }  # unresolved GUID stays visible - honest
        })
    }

    $policies = @($rawPolicies | Sort-Object displayName | ForEach-Object {
        $c = $_.conditions
        $grantControls = @()
        if ($_.grantControls.builtInControls) { $grantControls += @($_.grantControls.builtInControls) }
        if ($_.grantControls.authenticationStrength.displayName) {
            $grantControls += "Authentication strength: $($_.grantControls.authenticationStrength.displayName)"
        }
        $session = @()
        $s = $_.sessionControls
        if ($s.applicationEnforcedRestrictions.isEnabled) { $session += 'App-enforced restrictions' }
        if ($s.cloudAppSecurity.isEnabled) { $session += "Defender for Cloud Apps: $($s.cloudAppSecurity.cloudAppSecurityType)" }
        if ($s.signInFrequency.isEnabled) {
            $freq = if ($s.signInFrequency.frequencyInterval -eq 'everyTime') { 'every time' }
                    else { "$($s.signInFrequency.value) $($s.signInFrequency.type)" }
            $session += "Sign-in frequency: $freq"
        }
        if ($s.persistentBrowser.isEnabled) { $session += "Persistent browser: $($s.persistentBrowser.mode)" }

        [ordered]@{
            Name  = $_.displayName
            State = $_.state   # enabled | disabled | enabledForReportingButNotEnforced
            IncludeUsers  = @(Resolve-CaList $c.users.includeUsers  $dirNameById)
            ExcludeUsers  = @(Resolve-CaList $c.users.excludeUsers  $dirNameById)
            IncludeGroups = @(Resolve-CaList $c.users.includeGroups $dirNameById)
            ExcludeGroups = @(Resolve-CaList $c.users.excludeGroups $dirNameById)
            IncludeRoles  = @(Resolve-CaList $c.users.includeRoles  $roleNameById)
            ExcludeRoles  = @(Resolve-CaList $c.users.excludeRoles  $roleNameById)
            IncludeApps   = @(Resolve-CaList $c.applications.includeApplications @{} $WELL_KNOWN_APPS)
            ExcludeApps   = @(Resolve-CaList $c.applications.excludeApplications @{} $WELL_KNOWN_APPS)
            UserActions   = @($c.applications.includeUserActions)
            Platforms     = [ordered]@{
                Include = @($c.platforms.includePlatforms)
                Exclude = @($c.platforms.excludePlatforms)
            }
            Locations     = [ordered]@{
                Include = @(Resolve-CaList $c.locations.includeLocations $locNameById)
                Exclude = @(Resolve-CaList $c.locations.excludeLocations $locNameById)
            }
            ClientAppTypes  = @($c.clientAppTypes | Where-Object { $_ -ne 'all' })
            SignInRisk      = @($c.signInRiskLevels)
            UserRisk        = @($c.userRiskLevels)
            GrantOperator   = $_.grantControls.operator
            GrantControls   = $grantControls
            SessionControls = $session
        }
    })

    $namedLocations = @($rawLocations | Sort-Object displayName | ForEach-Object {
        $kind = if ($_.'@odata.type' -match 'ipNamedLocation') { 'IP ranges' }
                elseif ($_.'@odata.type' -match 'countryNamedLocation') { 'Countries' }
                else { 'Other' }
        $detail = if ($kind -eq 'IP ranges') {
            $trusted = if ($_.isTrusted) { 'trusted' } else { 'not trusted' }
            "$(@($_.ipRanges).Count) range(s), $trusted"
        } elseif ($kind -eq 'Countries') {
            (@($_.countriesAndRegions) -join ', ')
        } else { '' }
        [ordered]@{ Name = $_.displayName; Kind = $kind; Detail = $detail }
    })

    # -- 3. Role assignments ------------------------------------------------- #
    $assignments = Get-GraphPage "$G/roleManagement/directory/roleAssignments?`$expand=principal"
    $roleRows = @($assignments | ForEach-Object {
        [pscustomobject]@{
            Role = if ($roleNameById.ContainsKey($_.roleDefinitionId)) { $roleNameById[$_.roleDefinitionId] } else { $_.roleDefinitionId }
            PrincipalType = ($_.principal.'@odata.type' -replace '#microsoft.graph.', '')
            DisplayName = $_.principal.displayName
            UserPrincipalName = $_.principal.userPrincipalName
        }
    })
    $roles = @($roleRows | Group-Object Role | Sort-Object Name | ForEach-Object {
        [ordered]@{
            Role = $_.Name
            Members = @($_.Group | Sort-Object { $_.DisplayName } | ForEach-Object {
                [ordered]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Type = $_.PrincipalType }
            })
        }
    })

    # -- 4. Groups ------------------------------------------------------------ #
    Write-Host 'Collecting groups...'
    $rawGroups = Get-GraphPage ("$G/groups?`$select=id,displayName,groupTypes,securityEnabled,mailEnabled," +
        "isAssignableToRole,membershipRule,membershipRuleProcessingState&`$top=999")
    $dynamic = @($rawGroups | Where-Object { $_.groupTypes -contains 'DynamicMembership' } |
        Sort-Object displayName | ForEach-Object {
            [ordered]@{ Name = $_.displayName; Rule = $_.membershipRule; State = $_.membershipRuleProcessingState }
        })
    $groups = [ordered]@{
        Total          = @($rawGroups).Count
        M365           = @($rawGroups | Where-Object { $_.groupTypes -contains 'Unified' }).Count
        SecurityOnly   = @($rawGroups | Where-Object { $_.securityEnabled -and $_.groupTypes -notcontains 'Unified' }).Count
        MailEnabled    = @($rawGroups | Where-Object { $_.mailEnabled }).Count
        RoleAssignable = @($rawGroups | Where-Object { $_.isAssignableToRole } | Sort-Object displayName | ForEach-Object { $_.displayName })
        Dynamic        = $dynamic
    }

    # -- 5. Authentication methods policy ------------------------------------- #
    Write-Host 'Collecting authentication methods policy...'
    $authMethods = @()
    try {
        $amp = Invoke-MgGraphRequest -Method GET -Uri "$G/policies/authenticationMethodsPolicy" -OutputType PSObject
        $authMethods = @($amp.authenticationMethodConfigurations | Sort-Object id | ForEach-Object {
            [ordered]@{
                Method  = $_.id
                State   = $_.state
                Targets = @($_.includeTargets | ForEach-Object { $_.id })
            }
        })
    }
    catch { Write-Warning "Authentication methods policy unavailable ($($_.Exception.Message)) - section skipped." }

    # -- 6. Authorization policy (user & guest settings) ----------------------- #
    Write-Host 'Collecting authorization policy...'
    $userSettings = [ordered]@{}
    try {
        $ap = Invoke-MgGraphRequest -Method GET -Uri "$G/policies/authorizationPolicy" -OutputType PSObject
        $guestRoleNames = @{
            'a0b1b346-4d3e-4e8b-98f8-753987be4970' = 'Same as member users'
            '10dae51f-b6af-4016-8d66-8c2a99b929b3' = 'Limited access (default)'
            '2af84b1e-32c8-42b7-82bc-daa82404023b' = 'Restricted access'
        }
        $gr = $ap.guestUserRoleId
        $userSettings = [ordered]@{
            'Users can register applications'          = [bool]$ap.defaultUserRolePermissions.allowedToCreateApps
            'Users can create security groups'         = [bool]$ap.defaultUserRolePermissions.allowedToCreateSecurityGroups
            'Users can read other users'               = [bool]$ap.defaultUserRolePermissions.allowedToReadOtherUsers
            'Who can invite guests'                    = "$($ap.allowInvitesFrom)"
            'Guest user access level'                  = if ($gr -and $guestRoleNames.ContainsKey($gr)) { $guestRoleNames[$gr] } else { "$gr" }
            'Email-verified users can join the tenant' = [bool]$ap.allowEmailVerifiedUsersToJoinOrganization
            'Self-service password reset enabled'      = [bool]$ap.allowedToUseSSPR
        }
    }
    catch { Write-Warning "Authorization policy unavailable ($($_.Exception.Message)) - section skipped." }

    # -- 7. App registrations --------------------------------------------------- #
    Write-Host 'Collecting app registrations...'
    $rawApps = Get-GraphPage ("$G/applications?`$select=id,appId,displayName,signInAudience," +
        "createdDateTime,passwordCredentials,keyCredentials&`$top=999")
    $nowUtc = (Get-Date).ToUniversalTime()
    $apps = @($rawApps | Sort-Object displayName | ForEach-Object {
        $creds = @()
        foreach ($pc in @($_.passwordCredentials)) {
            $creds += [ordered]@{ Type = 'Client secret'; Name = $pc.displayName; ExpiresUtc = Iso $pc.endDateTime }
        }
        foreach ($kc in @($_.keyCredentials)) {
            $creds += [ordered]@{ Type = 'Certificate'; Name = $kc.displayName; ExpiresUtc = Iso $kc.endDateTime }
        }
        [ordered]@{
            Name           = $_.displayName
            AppId          = $_.appId
            SignInAudience = $_.signInAudience
            CreatedUtc     = Iso $_.createdDateTime
            Credentials    = $creds
        }
    })

    return [ordered]@{
        GeneratedUtc     = $nowUtc.ToString('o')
        TenantId         = (Get-MgContext).TenantId
        Organization     = [ordered]@{
            DisplayName = $org.displayName
            CreatedUtc  = Iso $org.createdDateTime
            Domains     = $domains
        }
        Licenses          = $licenses
        UserCounts        = $userCounts
        ConditionalAccess = [ordered]@{ Policies = $policies; NamedLocations = $namedLocations }
        Roles             = $roles
        Groups            = $groups
        AuthMethods       = $authMethods
        UserSettings      = $userSettings
        Applications      = $apps
    }
}

# --------------------------------------------------------------------------- #
# Shared render helpers
# --------------------------------------------------------------------------- #

function Get-CredRows {
    # Flatten app credentials with days-left computed against the SNAPSHOT time,
    # so re-rendering the same tenant.json is deterministic.
    param($Data, [int]$WarnDays)
    $anchor = [datetime]::Parse($Data.GeneratedUtc).ToUniversalTime()
    $rows = @(foreach ($app in @($Data.Applications)) {
        foreach ($c in @($app.Credentials)) {
            if (-not $c.ExpiresUtc) { continue }
            $days = [int][Math]::Floor(([datetime]::Parse($c.ExpiresUtc).ToUniversalTime() - $anchor).TotalDays)
            $sev = if ($days -lt 0) { 'expired' }        # renders with the critical color
                   elseif ($days -le 30) { 'serious' }
                   elseif ($days -le $WarnDays) { 'warning' }
                   else { 'ok' }
            [pscustomobject]@{
                App = $app.Name; Type = $c.Type; CredName = $c.Name
                ExpiresUtc = $c.ExpiresUtc; DaysLeft = $days; Severity = $sev
            }
        }
    })
    return @($rows | Sort-Object DaysLeft)
}

function Format-StateWord {
    param($State)
    switch ("$State") {
        'enabled'                             { 'Enabled' }
        'enabledForReportingButNotEnforced'   { 'Report-only' }
        'disabled'                            { 'Disabled' }
        default                               { "$State" }
    }
}

# --------------------------------------------------------------------------- #
# Markdown renderers - deterministic; timestamp appears ONLY in index.md
# --------------------------------------------------------------------------- #

function Write-Docs {
    param($Data, [string]$DocsPath, [int]$WarnDays)

    $null = New-Item -ItemType Directory -Path $DocsPath -Force
    $o = $Data.Organization
    $credRows = Get-CredRows $Data $WarnDays
    $expiring = @($credRows | Where-Object { $_.Severity -ne 'ok' })
    $caEnabled = @($Data.ConditionalAccess.Policies | Where-Object { $_.State -eq 'enabled' }).Count
    $caReport  = @($Data.ConditionalAccess.Policies | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' }).Count
    $caOff     = @($Data.ConditionalAccess.Policies | Where-Object { $_.State -eq 'disabled' }).Count

    # ---- index.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add("# Tenant documentation - $(MdEscape $o.DisplayName)")
    $md.Add('')
    $md.Add("Generated: **$($Data.GeneratedUtc)** (UTC) | Tenant ID: ``$($Data.TenantId)``")
    $md.Add('')
    $md.Add('Snapshot produced by [entra-tenant-docs](https://github.com/JonathanT10/entra-tenant-docs). Section files carry no timestamps, so a git diff of this folder is a config-drift report.')
    $md.Add('')
    $md.Add('| Section | At a glance |')
    $md.Add('|---|---|')
    $md.Add("| [1. Tenant](01-tenant.md) | $(@($o.Domains).Count) domain(s), $(@($Data.Licenses).Count) license SKU(s), $($Data.UserCounts.Members) members + $($Data.UserCounts.Guests) guests |")
    $md.Add("| [2. Conditional Access](02-conditional-access.md) | $caEnabled enabled, $caReport report-only, $caOff disabled |")
    $md.Add("| [3. Directory roles](03-roles.md) | $(@($Data.Roles).Count) role(s) with assignments |")
    $md.Add("| [4. Groups](04-groups.md) | $($Data.Groups.Total) total, $(@($Data.Groups.Dynamic).Count) dynamic |")
    $md.Add("| [5. Authentication methods](05-authentication.md) | $(@($Data.AuthMethods | Where-Object { $_.State -eq 'enabled' }).Count) of $(@($Data.AuthMethods).Count) methods enabled |")
    $md.Add("| [6. User & guest settings](06-user-settings.md) | $(@($Data.UserSettings.Keys).Count) settings documented |")
    $md.Add("| [7. Applications](07-applications.md) | $(@($Data.Applications).Count) registrations, $(@($expiring).Count) credential(s) expiring/expired |")
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath 'index.md') -Encoding UTF8

    # ---- 01-tenant.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add("# 1. Tenant")
    $md.Add('')
    $md.Add("| | |")
    $md.Add("|---|---|")
    $md.Add("| Name | $(MdEscape $o.DisplayName) |")
    $md.Add("| Tenant ID | ``$($Data.TenantId)`` |")
    $md.Add("| Created | $($o.CreatedUtc) |")
    $md.Add("| Members | $($Data.UserCounts.Members) ($($Data.UserCounts.EnabledMembers) enabled) |")
    $md.Add("| Guests | $($Data.UserCounts.Guests) |")
    $md.Add('')
    $md.Add('## Verified domains')
    $md.Add('')
    $md.Add('| Domain | Default | Initial |')
    $md.Add('|---|---|---|')
    foreach ($d in @($o.Domains)) {
        $md.Add("| $(MdEscape $d.Name) | $(if ($d.IsDefault) { 'yes' } else { '' }) | $(if ($d.IsInitial) { 'yes' } else { '' }) |")
    }
    $md.Add('')
    $md.Add('## License SKUs')
    $md.Add('')
    $md.Add('| SKU | Purchased | Assigned | Available |')
    $md.Add('|---|---:|---:|---:|')
    foreach ($l in @($Data.Licenses)) {
        $md.Add("| $(MdEscape $l.Sku) | $($l.Purchased) | $($l.Assigned) | $($l.Available) |")
    }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '01-tenant.md') -Encoding UTF8

    # ---- 02-conditional-access.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 2. Conditional Access')
    $md.Add('')
    $md.Add("$caEnabled enabled | $caReport report-only | $caOff disabled")
    foreach ($p in @($Data.ConditionalAccess.Policies)) {
        $md.Add('')
        $md.Add("## $(MdEscape $p.Name)")
        $md.Add('')
        $md.Add("**State: $(Format-StateWord $p.State)**")
        $md.Add('')
        $md.Add('| | |')
        $md.Add('|---|---|')
        $row = {
            param($label, $vals)
            $list = @(@($vals) | Where-Object { $null -ne $_ -and "$_" -ne '' })
            if ($list.Count) { $md.Add("| $label | $(MdEscape ($list -join '; ')) |") }
        }
        & $row 'Users included'  $p.IncludeUsers
        & $row 'Users excluded'  $p.ExcludeUsers
        & $row 'Groups included' $p.IncludeGroups
        & $row 'Groups excluded' $p.ExcludeGroups
        & $row 'Roles included'  $p.IncludeRoles
        & $row 'Roles excluded'  $p.ExcludeRoles
        & $row 'Apps included'   $p.IncludeApps
        & $row 'Apps excluded'   $p.ExcludeApps
        & $row 'User actions'    $p.UserActions
        & $row 'Platforms included' $p.Platforms.Include
        & $row 'Platforms excluded' $p.Platforms.Exclude
        & $row 'Locations included' $p.Locations.Include
        & $row 'Locations excluded' $p.Locations.Exclude
        & $row 'Client app types'   $p.ClientAppTypes
        & $row 'Sign-in risk'       $p.SignInRisk
        & $row 'User risk'          $p.UserRisk
        $grant = if (@($p.GrantControls).Count) {
            "$((@($p.GrantControls) -join "; ")) (operator: $($p.GrantOperator))"
        } else { '' }
        if ($grant) { $md.Add("| Grant | $(MdEscape $grant) |") }
        & $row 'Session' $p.SessionControls
    }
    $md.Add('')
    $md.Add('## Named locations')
    $md.Add('')
    if (@($Data.ConditionalAccess.NamedLocations).Count) {
        $md.Add('| Name | Kind | Detail |')
        $md.Add('|---|---|---|')
        foreach ($l in @($Data.ConditionalAccess.NamedLocations)) {
            $md.Add("| $(MdEscape $l.Name) | $($l.Kind) | $(MdEscape $l.Detail) |")
        }
    } else { $md.Add('None defined.') }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '02-conditional-access.md') -Encoding UTF8

    # ---- 03-roles.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 3. Directory roles')
    $md.Add('')
    $md.Add('Permanent assignments only - PIM eligible assignments are not included.')
    foreach ($r in @($Data.Roles)) {
        $md.Add('')
        $md.Add("## $(MdEscape $r.Role) ($(@($r.Members).Count))")
        $md.Add('')
        foreach ($m in @($r.Members)) {
            $upn = if ($m.UserPrincipalName) { " <$($m.UserPrincipalName)>" } else { '' }
            $md.Add("- $(MdEscape $m.DisplayName)$upn ($($m.Type))")
        }
    }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '03-roles.md') -Encoding UTF8

    # ---- 04-groups.md ---- #
    $g = $Data.Groups
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 4. Groups')
    $md.Add('')
    $md.Add('| | |')
    $md.Add('|---|---:|')
    $md.Add("| Total | $($g.Total) |")
    $md.Add("| Microsoft 365 | $($g.M365) |")
    $md.Add("| Security (non-M365) | $($g.SecurityOnly) |")
    $md.Add("| Mail-enabled | $($g.MailEnabled) |")
    $md.Add('')
    $md.Add('## Role-assignable groups')
    $md.Add('')
    if (@($g.RoleAssignable).Count) {
        foreach ($n in @($g.RoleAssignable)) { $md.Add("- $(MdEscape $n)") }
    } else { $md.Add('None.') }
    $md.Add('')
    $md.Add('## Dynamic groups and their rules')
    $md.Add('')
    if (@($g.Dynamic).Count) {
        foreach ($d in @($g.Dynamic)) {
            $md.Add("### $(MdEscape $d.Name)")
            $md.Add('')
            $md.Add("State: $($d.State)")
            $md.Add('')
            $md.Add('```')
            $md.Add("$($d.Rule)")
            $md.Add('```')
            $md.Add('')
        }
    } else { $md.Add('None.') }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '04-groups.md') -Encoding UTF8

    # ---- 05-authentication.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 5. Authentication methods policy')
    $md.Add('')
    if (@($Data.AuthMethods).Count) {
        $md.Add('| Method | State | Targets |')
        $md.Add('|---|---|---|')
        foreach ($m in @($Data.AuthMethods)) {
            $md.Add("| $(MdEscape $m.Method) | $($m.State) | $(MdEscape (@($m.Targets) -join '; ')) |")
        }
    } else { $md.Add('Not available in this snapshot.') }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '05-authentication.md') -Encoding UTF8

    # ---- 06-user-settings.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 6. User & guest settings')
    $md.Add('')
    if (@($Data.UserSettings.Keys).Count) {
        $md.Add('| Setting | Value |')
        $md.Add('|---|---|')
        foreach ($k in $Data.UserSettings.Keys) {
            $v = $Data.UserSettings[$k]
            $shown = if ($v -is [bool]) { if ($v) { 'Yes' } else { 'No' } } else { "$v" }
            $md.Add("| $(MdEscape $k) | $shown |")
        }
    } else { $md.Add('Not available in this snapshot.') }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '06-user-settings.md') -Encoding UTF8

    # ---- 07-applications.md ---- #
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 7. App registrations')
    $md.Add('')
    $md.Add("$(@($Data.Applications).Count) registration(s). Credential expiry is computed against the snapshot time.")
    $md.Add('')
    $md.Add('## Credentials needing attention')
    $md.Add('')
    $attention = @($credRows | Where-Object { $_.Severity -ne 'ok' })
    if (@($attention).Count) {
        $md.Add('| App | Credential | Expires (UTC) | Days left |')
        $md.Add('|---|---|---|---:|')
        foreach ($c in $attention) {
            $label = if ($c.DaysLeft -lt 0) { "EXPIRED ($($c.DaysLeft * -1) days ago)" } else { "$($c.DaysLeft)" }
            $md.Add("| $(MdEscape $c.App) | $($c.Type)$(if ($c.CredName) { " ($(MdEscape $c.CredName))" }) | $($c.ExpiresUtc) | $label |")
        }
    } else { $md.Add('None - nothing expiring inside the threshold.') }
    $md.Add('')
    $md.Add('## All registrations')
    $md.Add('')
    $md.Add('| App | Sign-in audience | Credentials | Created (UTC) |')
    $md.Add('|---|---|---:|---|')
    foreach ($a in @($Data.Applications)) {
        $md.Add("| $(MdEscape $a.Name) | $($a.SignInAudience) | $(@($a.Credentials).Count) | $($a.CreatedUtc) |")
    }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '07-applications.md') -Encoding UTF8
}

# --------------------------------------------------------------------------- #
# HTML report - the timestamped "face". Self-contained, no dependencies.
# --------------------------------------------------------------------------- #

$HTML_TEMPLATE = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Entra tenant report</title>
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb;
  --ink: #0b0b0b; --ink-2: #52514e; --muted: #898781;
  --grid: #e1e0d9; --baseline: #c3c2b7;
  --border: rgba(11,11,11,0.10);
  --accent: #2a78d6; --accent-deep: #1c5cab;
  --good: #0ca30c; --warning: #fab219; --serious: #ec835a; --critical: #d03b3b;
}
@media (prefers-color-scheme: dark) {
  :root:where(:not([data-theme="light"])) {
    color-scheme: dark;
    --page: #0d0d0d; --surface: #1a1a19;
    --ink: #ffffff; --ink-2: #c3c2b7; --muted: #898781;
    --grid: #2c2c2a; --baseline: #383835;
    --border: rgba(255,255,255,0.10);
    --accent: #3987e5; --accent-deep: #86b6ef;
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0d0d0d; --surface: #1a1a19;
  --ink: #ffffff; --ink-2: #c3c2b7; --muted: #898781;
  --grid: #2c2c2a; --baseline: #383835;
  --border: rgba(255,255,255,0.10);
  --accent: #3987e5; --accent-deep: #86b6ef;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--page); color: var(--ink);
  font: 14px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif; }
.wrap { max-width: 1120px; margin: 0 auto; padding: 24px 20px 48px; }
header { margin-bottom: 18px; }
header h1 { font-size: 21px; font-weight: 650; margin: 0 0 2px; }
header .sub { color: var(--ink-2); font-size: 13px; }
header .stamp { display: inline-block; margin-top: 8px; padding: 4px 10px;
  border: 1px solid var(--border); border-radius: 999px; background: var(--surface);
  font-size: 12.5px; color: var(--ink-2); }
header .stamp b { color: var(--ink); font-weight: 600; }
.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(165px, 1fr));
  gap: 12px; margin-bottom: 20px; }
.card { background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 12px 14px; }
.card .k { font-size: 12px; color: var(--ink-2); margin-bottom: 2px; }
.card .v { font-size: 22px; font-weight: 650; font-variant-numeric: tabular-nums; }
.card .d { font-size: 12px; color: var(--muted); }
section { background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 16px 18px; margin-bottom: 16px; }
section h2 { font-size: 15px; font-weight: 650; margin: 0 0 4px; }
section .note { font-size: 12.5px; color: var(--muted); margin: 0 0 10px; }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th { text-align: left; color: var(--ink-2); font-weight: 600; font-size: 12px;
  padding: 6px 10px 6px 0; border-bottom: 1px solid var(--baseline); }
td { padding: 7px 10px 7px 0; border-bottom: 1px solid var(--grid); vertical-align: top; }
tr:last-child td { border-bottom: none; }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
.badge { display: inline-flex; align-items: center; gap: 5px; font-size: 12px;
  font-weight: 600; white-space: nowrap; }
.badge svg { width: 12px; height: 12px; flex: none; }
.muted { color: var(--muted); }
.meter { display: flex; align-items: center; gap: 10px; }
.meter .track { flex: 1; height: 8px; border-radius: 4px; overflow: hidden;
  background: color-mix(in oklab, var(--accent) 16%, var(--surface)); }
.meter .fill { height: 100%; border-radius: 4px; background: var(--accent); }
.meter .lbl { font-size: 12px; color: var(--ink-2); font-variant-numeric: tabular-nums;
  white-space: nowrap; min-width: 86px; text-align: right; }
.bar-row { display: grid; grid-template-columns: minmax(180px, 260px) 1fr 44px;
  gap: 10px; align-items: center; padding: 5px 0; }
.bar-row .n { font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bar-row .bar { height: 12px; border-radius: 3px 4px 4px 3px; background: var(--accent); min-width: 2px; }
.bar-row .c { font-size: 12.5px; font-variant-numeric: tabular-nums; color: var(--ink-2); text-align: right; }
details { margin-top: 8px; }
summary { cursor: pointer; font-size: 12.5px; color: var(--accent-deep); }
details pre { background: var(--page); border: 1px solid var(--border); border-radius: 8px;
  padding: 10px 12px; font-size: 12px; overflow-x: auto; margin: 8px 0 0; }
ul.plain { margin: 6px 0 0; padding-left: 18px; font-size: 13px; }
ul.plain li { margin: 2px 0; }
footer { color: var(--muted); font-size: 12px; margin-top: 20px; }
.who { max-width: 340px; }
.who .x { color: var(--muted); font-size: 12px; }
@media (max-width: 700px) { .bar-row { grid-template-columns: 1fr 1fr 40px; } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1 id="t-name"></h1>
    <div class="sub">Entra ID tenant report &middot; read-only snapshot &middot; <span id="t-id"></span></div>
    <div class="stamp">Generated <b id="t-gen"></b> (UTC)</div>
  </header>
  <div class="cards" id="kpis"></div>
  <section><h2>Conditional Access</h2>
    <p class="note" id="ca-note"></p>
    <div style="overflow-x:auto"><table id="ca-table"></table></div>
    <div id="ca-locations"></div>
  </section>
  <section><h2>App credentials</h2>
    <p class="note" id="cred-note"></p>
    <div style="overflow-x:auto"><table id="cred-table"></table></div>
  </section>
  <section><h2>Directory role assignments</h2>
    <p class="note">Permanent assignments only &mdash; PIM eligible assignments are not included.</p>
    <div id="roles"></div>
  </section>
  <section><h2>Licenses</h2>
    <p class="note">Assigned vs purchased, per SKU.</p>
    <div id="licenses"></div>
  </section>
  <section><h2>Groups</h2>
    <div class="cards" id="group-tiles" style="margin-bottom:6px"></div>
    <div id="groups-extra"></div>
  </section>
  <section><h2>Authentication methods</h2>
    <div style="overflow-x:auto"><table id="auth-table"></table></div>
  </section>
  <section><h2>User &amp; guest settings</h2>
    <div style="overflow-x:auto"><table id="settings-table"></table></div>
  </section>
  <footer>Generated by <a href="https://github.com/JonathanT10/entra-tenant-docs">entra-tenant-docs</a>.
    Same snapshot as the Markdown docs and tenant.json beside this file.</footer>
</div>
<script>
const D = __PAYLOAD__;

const esc = s => String(s == null ? "" : s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const ICONS = {
  check: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M2.5 8.5l3.5 3.5 7-8"/></svg>',
  eye:   '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M1.5 8s2.5-4.5 6.5-4.5S14.5 8 14.5 8 12 12.5 8 12.5 1.5 8 1.5 8z"/><circle cx="8" cy="8" r="2"/></svg>',
  dash:  '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="8" r="6"/><path d="M5 8h6"/></svg>',
  warn:  '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M8 1.8L15 14H1L8 1.8z"/><path d="M8 6v4"/><circle cx="8" cy="12.2" r=".6" fill="currentColor"/></svg>',
  clock: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="8" cy="8" r="6"/><path d="M8 4.5V8l2.5 1.5"/></svg>'
};
function badge(color, icon, label) {
  return '<span class="badge" style="color:var(--' + color + ')">' + ICONS[icon] + esc(label) + '</span>';
}
const STATE_BADGE = {
  enabled: () => badge('good', 'check', 'Enabled'),
  enabledForReportingButNotEnforced: () => badge('warning', 'eye', 'Report-only'),
  disabled: () => '<span class="badge muted">' + ICONS.dash + 'Disabled</span>'
};

// ---- header ----
document.getElementById('t-name').textContent = D.Organization.DisplayName || 'Entra tenant';
document.getElementById('t-id').textContent = 'tenant ' + D.TenantId;
document.getElementById('t-gen').textContent = (D.GeneratedUtc || '').replace('T', ' ').replace(/(:\d\d)(\..*)?Z?$/, '$1');

// ---- credential rows (same thresholds as the docs: computed server-side severity) ----
const creds = (D.CredRows || []);
const credBad = creds.filter(c => c.Severity !== 'ok');

// ---- KPIs ----
const roleAssignTotal = (D.Roles || []).reduce((a, r) => a + (r.Members || []).length, 0);
const caPols = (D.ConditionalAccess.Policies || []);
const caEnabled = caPols.filter(p => p.State === 'enabled').length;
const kpis = [
  { k: 'Members', v: D.UserCounts.Members, d: D.UserCounts.EnabledMembers + ' enabled' },
  { k: 'Guests', v: D.UserCounts.Guests, d: '' },
  { k: 'CA policies enforced', v: caEnabled, d: 'of ' + caPols.length + ' total' },
  { k: 'Role assignments', v: roleAssignTotal, d: (D.Roles || []).length + ' roles in use' },
  { k: 'Groups', v: D.Groups.Total, d: (D.Groups.Dynamic || []).length + ' dynamic' },
  { k: 'Credentials to renew', v: credBad.length, d: creds.length + ' total tracked' }
];
document.getElementById('kpis').innerHTML = kpis.map(x =>
  '<div class="card"><div class="k">' + esc(x.k) + '</div><div class="v">' + esc(x.v) +
  '</div><div class="d">' + esc(x.d) + '</div></div>').join('');

// ---- Conditional Access ----
function whoSummary(p) {
  const parts = [];
  const inc = [];
  (p.IncludeUsers || []).forEach(u => inc.push(u === 'All' ? 'All users' : u));
  (p.IncludeGroups || []).forEach(g => inc.push(g + ' (group)'));
  (p.IncludeRoles || []).forEach(r => inc.push(r + ' (role)'));
  if (inc.length) {
    const shown = inc.slice(0, 3).join(', ');
    parts.push(shown + (inc.length > 3 ? ' +' + (inc.length - 3) + ' more' : ''));
  }
  const exc = [].concat(p.ExcludeUsers || [], p.ExcludeGroups || [], p.ExcludeRoles || []);
  let html = '<div class="who">' + esc(parts.join('; ') || '(none)');
  if (exc.length) {
    html += ' <span class="x" title="' + esc(exc.join(', ')) + '">&middot; excl. ' + exc.length + '</span>';
  }
  return html + '</div>';
}
function grantSummary(p) {
  const g = (p.GrantControls || []);
  if (!g.length) return '<span class="muted">(none)</span>';
  const op = g.length > 1 && p.GrantOperator ? ' <span class="muted">(' + esc(p.GrantOperator) + ')</span>' : '';
  return esc(g.join(', ')) + op;
}
document.getElementById('ca-note').textContent =
  caEnabled + ' enabled, ' +
  caPols.filter(p => p.State === 'enabledForReportingButNotEnforced').length + ' report-only, ' +
  caPols.filter(p => p.State === 'disabled').length + ' disabled. Full detail in docs/02-conditional-access.md.';
document.getElementById('ca-table').innerHTML =
  '<tr><th>Policy</th><th>State</th><th>Who</th><th>Apps</th><th>Grant</th></tr>' +
  caPols.map(p => {
    const apps = (p.IncludeApps || []).join(', ') || (p.UserActions || []).join(', ');
    const st = (STATE_BADGE[p.State] || (() => esc(p.State)))();
    return '<tr><td>' + esc(p.Name) + '</td><td>' + st + '</td><td>' + whoSummary(p) +
      '</td><td>' + esc(apps) + '</td><td>' + grantSummary(p) + '</td></tr>';
  }).join('');
const locs = (D.ConditionalAccess.NamedLocations || []);
if (locs.length) {
  document.getElementById('ca-locations').innerHTML =
    '<details><summary>Named locations (' + locs.length + ')</summary><ul class="plain">' +
    locs.map(l => '<li>' + esc(l.Name) + ' &mdash; ' + esc(l.Kind) +
      (l.Detail ? ' (' + esc(l.Detail) + ')' : '') + '</li>').join('') + '</ul></details>';
}

// ---- App credentials ----
const SEV = {
  expired: { color: 'critical', icon: 'warn', label: 'Expired' },
  serious: { color: 'serious', icon: 'clock', label: 'Expires soon' },
  warning: { color: 'warning', icon: 'clock', label: 'Expiring' }
};
document.getElementById('cred-note').textContent = credBad.length
  ? credBad.length + ' credential(s) inside the renewal window, of ' + creds.length + ' tracked.'
  : 'Nothing inside the renewal window. ' + creds.length + ' credential(s) tracked.';
document.getElementById('cred-table').innerHTML = credBad.length
  ? '<tr><th>App</th><th>Credential</th><th>Status</th><th class="num">Days left</th><th>Expires (UTC)</th></tr>' +
    credBad.map(c => {
      const s = SEV[c.Severity];
      return '<tr><td>' + esc(c.App) + '</td><td>' + esc(c.Type + (c.CredName ? ' (' + c.CredName + ')' : '')) +
        '</td><td>' + badge(s.color, s.icon, s.label) + '</td><td class="num">' +
        esc(c.DaysLeft) + '</td><td>' +
        esc((c.ExpiresUtc || '').slice(0, 10)) + '</td></tr>';
    }).join('')
  : '';

// ---- Roles (single-series bar list: direct labels, no legend) ----
const roles = (D.Roles || []);
const maxRole = Math.max(1, ...roles.map(r => (r.Members || []).length));
document.getElementById('roles').innerHTML = roles
  .slice().sort((a, b) => (b.Members || []).length - (a.Members || []).length)
  .map(r => {
    const n = (r.Members || []).length;
    return '<div class="bar-row"><div class="n" title="' + esc(r.Role) + '">' + esc(r.Role) +
      '</div><div><div class="bar" style="width:' + Math.round(100 * n / maxRole) +
      '%"></div></div><div class="c">' + n + '</div></div>';
  }).join('') +
  '<details><summary>Show every assignment</summary><ul class="plain">' +
  roles.map(r => (r.Members || []).map(m =>
    '<li>' + esc(r.Role) + ': ' + esc(m.DisplayName) +
    (m.UserPrincipalName ? ' &lt;' + esc(m.UserPrincipalName) + '&gt;' : '') +
    ' <span class="muted">(' + esc(m.Type) + ')</span></li>').join('')).join('') +
  '</ul></details>';

// ---- Licenses (meters) ----
document.getElementById('licenses').innerHTML = (D.Licenses || []).map(l => {
  const pct = l.Purchased ? Math.round(100 * l.Assigned / l.Purchased) : 0;
  return '<div class="bar-row"><div class="n" title="' + esc(l.Sku) + '">' + esc(l.Sku) +
    '</div><div class="meter"><div class="track"><div class="fill" style="width:' +
    Math.min(100, pct) + '%"></div></div><div class="lbl">' + l.Assigned + ' / ' + l.Purchased +
    '</div></div><div class="c">' + pct + '%</div></div>';
}).join('') || '<p class="note">No license data.</p>';

// ---- Groups ----
const g = D.Groups;
document.getElementById('group-tiles').innerHTML = [
  { k: 'Total', v: g.Total }, { k: 'Microsoft 365', v: g.M365 },
  { k: 'Security', v: g.SecurityOnly }, { k: 'Mail-enabled', v: g.MailEnabled },
  { k: 'Role-assignable', v: (g.RoleAssignable || []).length },
  { k: 'Dynamic', v: (g.Dynamic || []).length }
].map(x => '<div class="card"><div class="k">' + esc(x.k) + '</div><div class="v">' + esc(x.v) +
  '</div></div>').join('');
let gx = '';
if ((g.RoleAssignable || []).length) {
  gx += '<details><summary>Role-assignable groups</summary><ul class="plain">' +
    g.RoleAssignable.map(n => '<li>' + esc(n) + '</li>').join('') + '</ul></details>';
}
if ((g.Dynamic || []).length) {
  gx += '<details><summary>Dynamic membership rules</summary>' +
    g.Dynamic.map(d => '<p style="margin:8px 0 0;font-size:13px">' + esc(d.Name) +
      ' <span class="muted">(' + esc(d.State) + ')</span></p><pre>' + esc(d.Rule) + '</pre>').join('') +
    '</details>';
}
document.getElementById('groups-extra').innerHTML = gx;

// ---- Auth methods ----
document.getElementById('auth-table').innerHTML = (D.AuthMethods || []).length
  ? '<tr><th>Method</th><th>State</th><th>Targets</th></tr>' +
    D.AuthMethods.map(m => '<tr><td>' + esc(m.Method) + '</td><td>' +
      (m.State === 'enabled' ? badge('good', 'check', 'Enabled')
                             : '<span class="badge muted">' + ICONS.dash + 'Disabled</span>') +
      '</td><td>' + esc((m.Targets || []).join(', ')) + '</td></tr>').join('')
  : '<tr><td class="muted">Not available in this snapshot.</td></tr>';

// ---- Settings ----
const S = D.UserSettings || {};
document.getElementById('settings-table').innerHTML = Object.keys(S).length
  ? '<tr><th>Setting</th><th>Value</th></tr>' + Object.keys(S).map(k => {
      const v = S[k];
      const shown = v === true ? 'Yes' : v === false ? 'No' : String(v);
      return '<tr><td>' + esc(k) + '</td><td>' + esc(shown) + '</td></tr>';
    }).join('')
  : '<tr><td class="muted">Not available in this snapshot.</td></tr>';
</script>
</body>
</html>
'@

function Write-Report {
    param($Data, [string]$Path, [int]$WarnDays)
    # The HTML gets the same snapshot plus pre-computed credential rows, so the
    # page needs zero date math and stays consistent with the docs.
    $payload = [ordered]@{}
    foreach ($k in $Data.Keys) { $payload[$k] = $Data[$k] }
    $payload['CredRows'] = @(Get-CredRows $Data $WarnDays | ForEach-Object {
        [ordered]@{ App = $_.App; Type = $_.Type; CredName = $_.CredName
                    ExpiresUtc = $_.ExpiresUtc; DaysLeft = $_.DaysLeft; Severity = $_.Severity }
    })
    $json = ($payload | ConvertTo-Json -Depth 12 -Compress) -replace '</', '<\/'
    $HTML_TEMPLATE.Replace('__PAYLOAD__', $json) | Set-Content -Path $Path -Encoding UTF8
}

function Normalize-DataDates {
    # ConvertFrom-Json parses ISO strings into [datetime]; live collection can
    # also carry [datetime]. Canonicalize the four datetime fields to ISO
    # strings so every render path produces byte-identical output.
    param($Data)
    $Data['GeneratedUtc'] = Iso $Data.GeneratedUtc
    if ($Data.Organization) { $Data.Organization['CreatedUtc'] = Iso $Data.Organization.CreatedUtc }
    foreach ($a in @($Data.Applications)) {
        $a['CreatedUtc'] = Iso $a.CreatedUtc
        foreach ($c in @($a.Credentials)) { $c['ExpiresUtc'] = Iso $c.ExpiresUtc }
    }
}

# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

if ($SampleData) {
    $FromJson = Join-Path $PSScriptRoot 'sample-data.json'
}

$data = if ($FromJson) {
    if (-not (Test-Path $FromJson)) { exit "Input not found: $FromJson" }
    Write-Host "Rendering from $FromJson (no Graph connection)..."
    Get-Content $FromJson -Raw | ConvertFrom-Json -AsHashtable
} else {
    Get-TenantData -StaleCredDays $StaleCredDays
}

Normalize-DataDates $data
$null = New-Item -ItemType Directory -Path $OutputPath -Force
Write-Docs   -Data $data -DocsPath (Join-Path $OutputPath 'docs') -WarnDays $StaleCredDays
$data | ConvertTo-Json -Depth 12 | Set-Content -Path (Join-Path $OutputPath 'tenant.json') -Encoding UTF8
Write-Report -Data $data -Path (Join-Path $OutputPath 'report.html') -WarnDays $StaleCredDays

Write-Host ''
Write-Host "Done. Output in $OutputPath`:"
Write-Host '  docs/         8 Markdown files (commit these - the git diff is your drift report)'
Write-Host '  tenant.json   the full snapshot'
Write-Host '  report.html   the shareable report - open it in a browser'

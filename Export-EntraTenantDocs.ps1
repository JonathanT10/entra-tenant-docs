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
      2. Conditional Access - every policy rendered readable, named
         locations, and a 10-check gap analysis (MFA coverage, legacy auth,
         break-glass exclusions, lingering report-only policies, ...)
      3. Directory roles - permanent assignments per role
      4. Groups - counts, role-assignable groups, dynamic membership rules
      5. Authentication methods policy
      6. User & guest settings (authorization policy) in plain English
      7. App registrations - sign-in audience and credential expiry
      8. Intune - managed-device overview, compliance policies (settings +
         assignments), configuration profiles, app protection policies
      9. Change log - what changed between snapshots, computed by diffing

    Each run also archives its snapshot to history/ - from two snapshots
    onward the report gains trend sparklines and a "what changed" feed -
    and writes run-summary.json (this run's delta), which the companion
    Send-TenantDocsAlert.ps1 turns into Teams/email alerts.

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

.PARAMETER HistoryPath
    Folder where each run archives its snapshot JSON, and where trends and
    the change log are computed from. Default: <OutputPath>\history
    (with -SampleData: the bundled sample-history folder).

.PARAMETER NoHistory
    Do not archive this run's snapshot. Trends and the change log still
    render from whatever history already exists.

.PARAMETER SkipIntune
    Skip the Intune section entirely (tenants without Intune degrade
    automatically; this just avoids the attempt).

.PARAMETER Anonymize
    Replace identifying values in the OUTPUT with stable pseudonyms, so the
    docs, report and JSON can be shown to someone outside the tenant.

    Replaced: tenant and app ids, organization name, domains, every person's
    display name and UPN, group names, non-Microsoft app names, named-location
    names, app-registration and credential names, Intune assignment groups, and
    the string literals inside dynamic membership rules.

    Kept, because the shape is the point: all counts, license SKU part numbers,
    built-in role names, policy states, grant and session controls, platforms,
    client app types, Intune types and setting keys, and dates.

    ALSO KEPT, and worth knowing before you share: Conditional Access and
    Intune POLICY NAMES are left as written, so the gap analysis stays
    readable. Those are admin-authored free text and can name a vendor, a
    person, or a service account. Read them before handing the output over.

    The history archived on disk is NOT anonymized - only the rendered output
    is - so your own drift history stays intact and readable.

.PARAMETER AnonymizeSalt
    Fixes the anonymization salt so the same real value maps to the same
    pseudonym on every run. Without it a fresh random salt is used per run,
    which is the safer default: anyone holding the output cannot confirm a
    guessed name by re-deriving its pseudonym. Use this only when you need
    two separately generated reports to line up.

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
        + for Intune: DeviceManagementConfiguration.Read.All,
          DeviceManagementManagedDevices.Read.All, DeviceManagementApps.Read.All
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
    [ValidateRange(1, 3650)][int]$StaleCredDays = 90,
    [string]$HistoryPath,
    [switch]$NoHistory,
    [switch]$SkipIntune,
    [switch]$Anonymize,
    [string]$AnonymizeSalt
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
    param([int]$StaleCredDays, [switch]$SkipIntune)

    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'Directory.Read.All','Policy.Read.All',
            'RoleManagement.Read.Directory','Application.Read.All',
            'DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.Read.All',
            'DeviceManagementApps.Read.All' -NoWelcome
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

    $securityDefaults = $null
    try {
        $sd = Invoke-MgGraphRequest -Method GET -Uri "$G/policies/identitySecurityDefaultsEnforcementPolicy" -OutputType PSObject
        $securityDefaults = [bool]$sd.isEnabled
    }
    catch { Write-Warning "Security defaults state unavailable ($($_.Exception.Message))." }

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

    # -- 8. Intune (v1.0 endpoints only; degrades when absent) ----------------- #
    $intune = [ordered]@{ Available = $false }
    if (-not $SkipIntune) {
        Write-Host 'Collecting Intune (skipped automatically if not licensed)...'
        try {
            $ov = Invoke-MgGraphRequest -Method GET -Uri "$G/deviceManagement/managedDeviceOverview" -OutputType PSObject
            $os = $ov.deviceOperatingSystemSummary
            $devices = [ordered]@{
                Total   = [int]$ov.enrolledDeviceCount
                Windows = [int]$os.windowsCount
                MacOS   = [int]$os.macOSCount
                IOS     = [int]$os.iosCount
                Android = [int]$os.androidCount
                Other   = [Math]::Max(0, [int]$ov.enrolledDeviceCount - [int]$os.windowsCount - [int]$os.macOSCount - [int]$os.iosCount - [int]$os.androidCount)
            }
            $cs = Invoke-MgGraphRequest -Method GET -Uri "$G/deviceManagement/deviceCompliancePolicyDeviceStateSummary" -OutputType PSObject
            $complianceSummary = [ordered]@{
                Compliant     = [int]$cs.compliantDeviceCount
                NonCompliant  = [int]$cs.nonCompliantDeviceCount
                InGracePeriod = [int]$cs.inGracePeriodCount
                Error         = [int]$cs.errorDeviceCount
                Conflict      = [int]$cs.conflictDeviceCount
                Unknown       = [int]$cs.unknownDeviceCount
            }

            $rawCompliance = Get-GraphPage "$G/deviceManagement/deviceCompliancePolicies?`$expand=assignments"
            $rawConfigs    = Get-GraphPage "$G/deviceManagement/deviceConfigurations?`$expand=assignments"

            # Resolve assignment group ids to names (one POST)
            $iIds = New-Object System.Collections.Generic.HashSet[string]
            foreach ($pol in @($rawCompliance) + @($rawConfigs)) {
                foreach ($a in @($pol.assignments)) {
                    if ($a.target.groupId) { [void]$iIds.Add("$($a.target.groupId)") }
                }
            }
            $iNameById = @{}
            if ($iIds.Count) {
                try {
                    $b = @{ ids = @($iIds); types = @('group') } | ConvertTo-Json
                    $res = Invoke-MgGraphRequest -Method POST -Uri "$G/directoryObjects/getByIds" `
                        -Body $b -ContentType 'application/json' -OutputType PSObject
                    foreach ($o in @($res.value)) { $iNameById[$o.id] = $o.displayName }
                }
                catch { Write-Warning "Intune getByIds failed ($($_.Exception.Message)) - group GUIDs left unresolved." }
            }
            function Resolve-IntuneAssignments {
                param($Assignments)
                $out = @(foreach ($a in @($Assignments)) {
                    $t = $a.target
                    switch -Wildcard ("$($t.'@odata.type')") {
                        '*allDevicesAssignmentTarget'        { 'All devices' }
                        '*allLicensedUsersAssignmentTarget'  { 'All users' }
                        '*exclusionGroupAssignmentTarget'    {
                            $n = if ($iNameById.ContainsKey("$($t.groupId)")) { $iNameById["$($t.groupId)"] } else { "$($t.groupId)" }
                            "Exclude: $n" }
                        '*groupAssignmentTarget'             {
                            if ($iNameById.ContainsKey("$($t.groupId)")) { $iNameById["$($t.groupId)"] } else { "$($t.groupId)" } }
                        default                              { "$($t.'@odata.type')" -replace '#microsoft.graph.', '' }
                    }
                })
                return @($out | Sort-Object)
            }
            function Get-FriendlyType {
                param($OdataType)
                ("$OdataType" -replace '#microsoft.graph.', '')
            }
            # Generic scalar-settings dump: honest and vendor-agnostic
            $SKIP_PROPS = @('id', 'displayName', 'description', 'createdDateTime', 'lastModifiedDateTime',
                            'version', '@odata.type', 'roleScopeTagIds', 'assignments', 'scheduledActionsForRule',
                            'deviceSettingStateSummaries', 'deviceStatuses', 'userStatuses',
                            'deviceStatusOverview', 'userStatusOverview')
            $compliancePolicies = @($rawCompliance | Sort-Object displayName | ForEach-Object {
                $settings = [ordered]@{}
                foreach ($prop in ($_.PSObject.Properties | Sort-Object Name)) {
                    if ($prop.Name -in $SKIP_PROPS) { continue }
                    $v = $prop.Value
                    if ($null -eq $v) { continue }
                    if ($v -is [datetime]) { $settings[$prop.Name] = Iso $v }
                    elseif ($v -is [bool] -or $v -is [string] -or $v -is [int] -or $v -is [long] -or $v -is [double]) {
                        $settings[$prop.Name] = $v
                    }
                }
                [ordered]@{
                    Name        = $_.displayName
                    Type        = Get-FriendlyType $_.'@odata.type'
                    Assignments = Resolve-IntuneAssignments $_.assignments
                    Settings    = $settings
                }
            })
            $configurationProfiles = @($rawConfigs | Sort-Object displayName | ForEach-Object {
                [ordered]@{
                    Name        = $_.displayName
                    Type        = Get-FriendlyType $_.'@odata.type'
                    Assignments = Resolve-IntuneAssignments $_.assignments
                }
            })

            $appProtection = @()
            try {
                $rawMam = Get-GraphPage "$G/deviceAppManagement/managedAppPolicies"
                $appProtection = @($rawMam | Sort-Object displayName | ForEach-Object {
                    [ordered]@{ Name = $_.displayName; Type = Get-FriendlyType $_.'@odata.type' }
                })
            }
            catch { Write-Warning "App protection policies unavailable ($($_.Exception.Message)) - subsection skipped." }

            $intune = [ordered]@{
                Available             = $true
                Devices               = $devices
                ComplianceSummary     = $complianceSummary
                CompliancePolicies    = $compliancePolicies
                ConfigurationProfiles = $configurationProfiles
                AppProtection         = $appProtection
            }
        }
        catch {
            Write-Warning "Intune unavailable ($($_.Exception.Message)) - section skipped. (Needs an Intune license and the DeviceManagement*.Read.All scopes.)"
        }
    }

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
        ConditionalAccess = [ordered]@{ Policies = $policies; NamedLocations = $namedLocations; SecurityDefaults = $securityDefaults }
        Roles             = $roles
        Groups            = $groups
        AuthMethods       = $authMethods
        UserSettings      = $userSettings
        Applications      = $apps
        Intune            = $intune
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
# Conditional Access gap analysis - opinionated baseline hygiene checks,
# computed purely from the snapshot (works offline and on old snapshots).
# --------------------------------------------------------------------------- #

function Get-CaGapAnalysis {
    param($Data)
    $pols = @($Data.ConditionalAccess.Policies)
    $enabled = @($pols | Where-Object { "$($_.State)" -eq 'enabled' })
    $reportOnly = @($pols | Where-Object { "$($_.State)" -eq 'enabledForReportingButNotEnforced' })

    function _allUsers($p)  { @($p.IncludeUsers) -contains 'All' }
    function _allApps($p)   { @($p.IncludeApps) -contains 'All cloud apps' }
    function _mfa($p) {
        (@($p.GrantControls) -contains 'mfa') -or
        (@(@($p.GrantControls) | Where-Object { "$_" -like 'Authentication strength:*' }).Count -gt 0)
    }
    function _block($p)     { @($p.GrantControls) -contains 'block' }
    function _legacy($p)    { (@($p.ClientAppTypes) -contains 'exchangeActiveSync') -or (@($p.ClientAppTypes) -contains 'other') }

    $checks = New-Object System.Collections.Generic.List[object]
    $chk = { param($id, $title, $sev, $result, $detail)
        $checks.Add([ordered]@{ Id = $id; Title = $title; Severity = $sev; Result = $result; Detail = "$detail" })
    }

    # C1: MFA (or auth strength) required for all users, all apps
    $c1 = @($enabled | Where-Object { (_allUsers $_) -and (_allApps $_) -and (_mfa $_) })
    if ($c1.Count) { & $chk 'mfa-all-users' 'MFA required for all users' 'critical' 'pass' "satisfied by '$($c1[0].Name)'" }
    else { & $chk 'mfa-all-users' 'MFA required for all users' 'critical' 'fail' 'no enabled policy requires MFA for all users on all apps' }

    # C2: legacy authentication blocked
    $c2 = @($enabled | Where-Object { (_block $_) -and (_legacy $_) })
    if ($c2.Count) { & $chk 'block-legacy-auth' 'Legacy authentication blocked' 'critical' 'pass' "satisfied by '$($c2[0].Name)'" }
    else { & $chk 'block-legacy-auth' 'Legacy authentication blocked' 'critical' 'fail' 'no enabled policy blocks legacy client apps (Exchange ActiveSync / other)' }

    # C3: admins covered by MFA (role-targeted, or the all-users policy)
    $c3role = @($enabled | Where-Object { @($_.IncludeRoles).Count -gt 0 -and (_mfa $_) })
    if ($c3role.Count) { & $chk 'admin-mfa' 'Admin roles require MFA' 'critical' 'pass' "role-targeted: '$($c3role[0].Name)'" }
    elseif ($c1.Count) { & $chk 'admin-mfa' 'Admin roles require MFA' 'critical' 'pass' "covered by the all-users policy '$($c1[0].Name)'" }
    else {
        $c3ro = @($reportOnly | Where-Object { @($_.IncludeRoles).Count -gt 0 -and (_mfa $_) })
        $d = if ($c3ro.Count) { "only report-only coverage ('$($c3ro[0].Name)') - not enforced" } else { 'no enabled policy targets directory roles with MFA' }
        & $chk 'admin-mfa' 'Admin roles require MFA' 'critical' 'fail' $d
    }

    # C4: some baseline exists at all
    $sdState = if ($Data.ConditionalAccess.Contains('SecurityDefaults')) { $Data.ConditionalAccess.SecurityDefaults } else { $null }
    if ($enabled.Count -gt 0) { & $chk 'baseline-exists' 'Conditional Access is enforced' 'critical' 'pass' "$($enabled.Count) enabled policies" }
    elseif ($sdState -eq $true) { & $chk 'baseline-exists' 'Conditional Access is enforced' 'critical' 'pass' 'no CA policies, but security defaults are on' }
    elseif ($null -eq $sdState) { & $chk 'baseline-exists' 'Conditional Access is enforced' 'critical' 'unknown' 'no enabled CA policies; security defaults state unknown in this snapshot' }
    else { & $chk 'baseline-exists' 'Conditional Access is enforced' 'critical' 'fail' 'no enabled CA policies AND security defaults are off' }

    # C5: break-glass exclusions on lockout-capable all-users policies
    $c5off = @($enabled | Where-Object { (_allUsers $_) -and ((_block $_) -or (_mfa $_)) -and
        ((@($_.ExcludeUsers).Count + @($_.ExcludeGroups).Count) -eq 0) })
    if ($c5off.Count) { & $chk 'breakglass-exclusion' 'Break-glass accounts excluded' 'warning' 'fail' "no exclusions on: $((@($c5off | ForEach-Object { $_.Name }) -join ', ')) - lockout risk" }
    else { & $chk 'breakglass-exclusion' 'Break-glass accounts excluded' 'warning' 'pass' 'every lockout-capable all-users policy has exclusions' }

    # C6: guests covered
    $c6 = @($enabled | Where-Object { ((_mfa $_) -or (_block $_)) -and
        ((@($_.IncludeUsers) -contains 'GuestsOrExternalUsers') -or (_allUsers $_)) })
    if ($c6.Count) { & $chk 'guest-protection' 'Guests covered by MFA or block' 'warning' 'pass' "satisfied by '$($c6[0].Name)'" }
    else { & $chk 'guest-protection' 'Guests covered by MFA or block' 'warning' 'fail' 'no enabled policy covers guests' }

    # C7: risk-based policies (Entra ID P2)
    $c7 = @($enabled | Where-Object { @($_.SignInRisk).Count -gt 0 -or @($_.UserRisk).Count -gt 0 })
    if ($c7.Count) { & $chk 'risk-policies' 'Risk-based policies in use' 'info' 'pass' "satisfied by '$($c7[0].Name)'" }
    else { & $chk 'risk-policies' 'Risk-based policies in use' 'info' 'fail' 'no enabled policy uses sign-in or user risk (Entra ID P2 feature)' }

    # C8: device-based grant somewhere
    $c8 = @($enabled | Where-Object { (@($_.GrantControls) -contains 'compliantDevice') -or (@($_.GrantControls) -contains 'domainJoinedDevice') })
    if ($c8.Count) { & $chk 'device-grants' 'Device compliance used in grants' 'info' 'pass' "satisfied by '$($c8[0].Name)'" }
    else { & $chk 'device-grants' 'Device compliance used in grants' 'info' 'fail' 'no enabled policy requires a compliant or hybrid-joined device' }

    # C9: report-only policies lingering
    if ($reportOnly.Count) { & $chk 'report-only-lingering' 'No lingering report-only policies' 'warning' 'fail' "report-only: $((@($reportOnly | ForEach-Object { $_.Name }) -join ', ')) - enforce or remove" }
    else { & $chk 'report-only-lingering' 'No lingering report-only policies' 'warning' 'pass' '' }

    # C10: named locations all referenced
    $usedLoc = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in $pols) {
        foreach ($n in @($p.Locations.Include) + @($p.Locations.Exclude)) { if ($n) { [void]$usedLoc.Add("$n") } }
    }
    $unused = @(@($Data.ConditionalAccess.NamedLocations) | Where-Object { -not $usedLoc.Contains("$($_.Name)") } | ForEach-Object { $_.Name })
    if ($unused.Count) { & $chk 'unused-locations' 'Named locations all referenced' 'info' 'fail' "unused: $($unused -join ', ')" }
    else { & $chk 'unused-locations' 'Named locations all referenced' 'info' 'pass' '' }

    return $checks.ToArray()
}

# --------------------------------------------------------------------------- #
# History store, trend series, change detection
# --------------------------------------------------------------------------- #

function Save-Snapshot {
    param($Data, [string]$HistoryDir)
    $null = New-Item -ItemType Directory -Path $HistoryDir -Force
    $stamp = ([datetime]::Parse($Data.GeneratedUtc).ToUniversalTime()).ToString('yyyyMMdd-HHmmss')
    $path = Join-Path $HistoryDir "snapshot-$stamp.json"
    $Data | ConvertTo-Json -Depth 12 | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Get-History {
    # Every archived snapshot plus the current one, deduped by GeneratedUtc,
    # oldest first. Old snapshots may predate newer fields - readers must
    # tolerate missing keys.
    param([string]$HistoryDir, $Current)
    $snaps = @()
    if ($HistoryDir -and (Test-Path $HistoryDir)) {
        foreach ($f in (Get-ChildItem $HistoryDir -Filter '*.json' | Sort-Object Name)) {
            try {
                $loaded = Get-Content $f.FullName -Raw | ConvertFrom-Json -AsHashtable
                Normalize-DataDates $loaded
                $snaps += , $loaded
            }
            catch { Write-Warning "Skipping unreadable history file: $($f.Name)" }
        }
    }
    $seen = @{}; $out = @()
    foreach ($s in (@($snaps) + @($Current) | Sort-Object { "$($_.GeneratedUtc)" })) {
        $k = "$($s.GeneratedUtc)"
        if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $out += , $s }
    }
    return , $out
}

function Get-TrendSeries {
    # One metrics row per snapshot, oldest first. Pure function of the data.
    param($Snapshots, [int]$WarnDays)
    $rows = @(foreach ($s in $Snapshots) {
        $ra = 0; $ga = 0
        foreach ($r in @($s.Roles)) {
            $n = @($r.Members).Count
            $ra += $n
            if ("$($r.Role)" -eq 'Global Administrator') { $ga += $n }
        }
        $la = 0; foreach ($l in @($s.Licenses)) { $la += [int]$l.Assigned }
        [ordered]@{
            Ts               = "$($s.GeneratedUtc)"
            Members          = [int]$s.UserCounts.Members
            EnabledMembers   = [int]$s.UserCounts.EnabledMembers
            Guests           = [int]$s.UserCounts.Guests
            CaEnabled        = @(@($s.ConditionalAccess.Policies) | Where-Object { $_.State -eq 'enabled' }).Count
            CaTotal          = @($s.ConditionalAccess.Policies).Count
            RoleAssignments  = $ra
            GlobalAdmins     = $ga
            Groups           = [int]$s.Groups.Total
            AppRegistrations = @($s.Applications).Count
            LicenseAssigned  = $la
            CredsInWindow    = @(Get-CredRows $s $WarnDays | Where-Object { $_.Severity -ne 'ok' }).Count
            CaGapsCritical   = @(Get-CaGapAnalysis $s | Where-Object { $_.Result -eq 'fail' -and $_.Severity -eq 'critical' }).Count
            EnrolledDevices  = if ($s.Intune -and $s.Intune.Available) { [int]$s.Intune.Devices.Total } else { $null }
            NonCompliantDevices = if ($s.Intune -and $s.Intune.Available) { [int]$s.Intune.ComplianceSummary.NonCompliant } else { $null }
        }
    })
    return , $rows
}

function Get-SnapshotChanges {
    # Human-readable diff between two consecutive snapshots.
    param($Prev, $Curr)
    $L = New-Object System.Collections.Generic.List[object]
    $ts = "$($Curr.GeneratedUtc)"
    $add = { param($cat, $kind, $item, $detail)
        $L.Add([ordered]@{ Ts = $ts; Category = $cat; Kind = $kind; Item = "$item"; Detail = "$detail" })
    }
    function _map($rows, $keyExpr) {
        $m = [ordered]@{}
        foreach ($r in @($rows)) { if ($null -ne $r) { $m["$(& $keyExpr $r)"] = $r } }
        return $m
    }

    # Conditional Access policies (by name)
    $pp = _map $Prev.ConditionalAccess.Policies { param($x) $x.Name }
    $cp = _map $Curr.ConditionalAccess.Policies { param($x) $x.Name }
    foreach ($k in $cp.Keys) {
        if (-not $pp.Contains($k)) { & $add 'Conditional Access' 'added' $k "state: $(Format-StateWord $cp[$k].State)"; continue }
        $a = $pp[$k]; $b = $cp[$k]
        if ("$($a.State)" -ne "$($b.State)") {
            & $add 'Conditional Access' 'changed' $k "$(Format-StateWord $a.State) -> $(Format-StateWord $b.State)"
        }
        elseif (($a | ConvertTo-Json -Depth 8 -Compress) -ne ($b | ConvertTo-Json -Depth 8 -Compress)) {
            & $add 'Conditional Access' 'changed' $k 'definition changed'
        }
    }
    foreach ($k in $pp.Keys) { if (-not $cp.Contains($k)) { & $add 'Conditional Access' 'removed' $k '' } }

    # Named locations (by name)
    $pl = _map $Prev.ConditionalAccess.NamedLocations { param($x) $x.Name }
    $cl = _map $Curr.ConditionalAccess.NamedLocations { param($x) $x.Name }
    foreach ($k in $cl.Keys) {
        if (-not $pl.Contains($k)) { & $add 'Named locations' 'added' $k $cl[$k].Detail }
        elseif ("$($pl[$k].Detail)" -ne "$($cl[$k].Detail)") { & $add 'Named locations' 'changed' $k "$($pl[$k].Detail) -> $($cl[$k].Detail)" }
    }
    foreach ($k in $pl.Keys) { if (-not $cl.Contains($k)) { & $add 'Named locations' 'removed' $k '' } }

    # Role assignments (flattened role|principal)
    function _roleSet($snap) {
        $set = [ordered]@{}
        foreach ($r in @($snap.Roles)) {
            foreach ($m in @($r.Members)) {
                $who = if ($m.UserPrincipalName) { "$($m.DisplayName) <$($m.UserPrincipalName)>" } else { "$($m.DisplayName) ($($m.Type))" }
                $set["$($r.Role)|$who"] = @($r.Role, $who)
            }
        }
        return $set
    }
    $pr = _roleSet $Prev; $cr = _roleSet $Curr
    foreach ($k in $cr.Keys) { if (-not $pr.Contains($k)) { & $add 'Role assignments' 'added' $cr[$k][1] $cr[$k][0] } }
    foreach ($k in $pr.Keys) { if (-not $cr.Contains($k)) { & $add 'Role assignments' 'removed' $pr[$k][1] $pr[$k][0] } }

    # Licenses (purchased counts and SKU add/remove; assigned drift lives in trends)
    $ps = _map $Prev.Licenses { param($x) $x.Sku }
    $cs = _map $Curr.Licenses { param($x) $x.Sku }
    foreach ($k in $cs.Keys) {
        if (-not $ps.Contains($k)) { & $add 'Licenses' 'added' $k "$($cs[$k].Purchased) purchased" }
        elseif ([int]$ps[$k].Purchased -ne [int]$cs[$k].Purchased) { & $add 'Licenses' 'changed' $k "purchased $($ps[$k].Purchased) -> $($cs[$k].Purchased)" }
    }
    foreach ($k in $ps.Keys) { if (-not $cs.Contains($k)) { & $add 'Licenses' 'removed' $k '' } }

    # Dynamic groups (by name): rule or state changes
    $pd = _map $Prev.Groups.Dynamic { param($x) $x.Name }
    $cd = _map $Curr.Groups.Dynamic { param($x) $x.Name }
    foreach ($k in $cd.Keys) {
        if (-not $pd.Contains($k)) { & $add 'Dynamic groups' 'added' $k ''; continue }
        if ("$($pd[$k].Rule)" -ne "$($cd[$k].Rule)") { & $add 'Dynamic groups' 'changed' $k 'membership rule changed' }
        elseif ("$($pd[$k].State)" -ne "$($cd[$k].State)") { & $add 'Dynamic groups' 'changed' $k "processing $($pd[$k].State) -> $($cd[$k].State)" }
    }
    foreach ($k in $pd.Keys) { if (-not $cd.Contains($k)) { & $add 'Dynamic groups' 'removed' $k '' } }

    # Role-assignable groups (name lists)
    $pra = @($Prev.Groups.RoleAssignable); $cra = @($Curr.Groups.RoleAssignable)
    foreach ($n in $cra) { if ($n -notin $pra) { & $add 'Role-assignable groups' 'added' $n '' } }
    foreach ($n in $pra) { if ($n -notin $cra) { & $add 'Role-assignable groups' 'removed' $n '' } }

    # Authentication methods (by method): state changes
    $pa = _map $Prev.AuthMethods { param($x) $x.Method }
    $ca = _map $Curr.AuthMethods { param($x) $x.Method }
    foreach ($k in $ca.Keys) {
        if ($pa.Contains($k) -and "$($pa[$k].State)" -ne "$($ca[$k].State)") {
            & $add 'Authentication methods' 'changed' $k "$($pa[$k].State) -> $($ca[$k].State)"
        }
    }

    # User & guest settings (by key)
    function _fmtVal($v) { if ($v -is [bool]) { if ($v) { 'Yes' } else { 'No' } } else { "$v" } }
    if ($Prev.UserSettings -and $Curr.UserSettings) {
        foreach ($k in $Curr.UserSettings.Keys) {
            if ($Prev.UserSettings.Contains($k) -and "$($Prev.UserSettings[$k])" -ne "$($Curr.UserSettings[$k])") {
                & $add 'User settings' 'changed' $k "$(_fmtVal $Prev.UserSettings[$k]) -> $(_fmtVal $Curr.UserSettings[$k])"
            }
        }
    }

    # Intune (only when both snapshots have it - older snapshots may predate the section)
    if ($Prev.Intune -and $Curr.Intune -and $Prev.Intune.Available -and $Curr.Intune.Available) {
        $pc = _map $Prev.Intune.CompliancePolicies { param($x) $x.Name }
        $cc = _map $Curr.Intune.CompliancePolicies { param($x) $x.Name }
        foreach ($k in $cc.Keys) {
            if (-not $pc.Contains($k)) { & $add 'Intune compliance' 'added' $k $cc[$k].Type; continue }
            if (($pc[$k] | ConvertTo-Json -Depth 6 -Compress) -ne ($cc[$k] | ConvertTo-Json -Depth 6 -Compress)) {
                & $add 'Intune compliance' 'changed' $k 'settings or assignments changed'
            }
        }
        foreach ($k in $pc.Keys) { if (-not $cc.Contains($k)) { & $add 'Intune compliance' 'removed' $k '' } }

        $pf = _map $Prev.Intune.ConfigurationProfiles { param($x) $x.Name }
        $cf = _map $Curr.Intune.ConfigurationProfiles { param($x) $x.Name }
        foreach ($k in $cf.Keys) {
            if (-not $pf.Contains($k)) { & $add 'Intune configuration' 'added' $k $cf[$k].Type; continue }
            if (($pf[$k] | ConvertTo-Json -Depth 6 -Compress) -ne ($cf[$k] | ConvertTo-Json -Depth 6 -Compress)) {
                & $add 'Intune configuration' 'changed' $k 'type or assignments changed'
            }
        }
        foreach ($k in $pf.Keys) { if (-not $cf.Contains($k)) { & $add 'Intune configuration' 'removed' $k '' } }

        $pm = _map $Prev.Intune.AppProtection { param($x) $x.Name }
        $cm = _map $Curr.Intune.AppProtection { param($x) $x.Name }
        foreach ($k in $cm.Keys) { if (-not $pm.Contains($k)) { & $add 'App protection' 'added' $k '' } }
        foreach ($k in $pm.Keys) { if (-not $cm.Contains($k)) { & $add 'App protection' 'removed' $k '' } }
    }

    # CA gap transitions (derived posture drift - opens and closes alert-worthy)
    $pg = _map (Get-CaGapAnalysis $Prev) { param($x) $x.Id }
    $cg = _map (Get-CaGapAnalysis $Curr) { param($x) $x.Id }
    foreach ($k in $cg.Keys) {
        if (-not $pg.Contains($k)) { continue }
        $a = $pg[$k]; $b = $cg[$k]
        if ("$($a.Result)" -eq 'unknown' -or "$($b.Result)" -eq 'unknown') { continue }
        if ("$($a.Result)" -eq 'pass' -and "$($b.Result)" -eq 'fail') {
            & $add 'CA gap' 'added' $b.Title "now failing ($($b.Severity)): $($b.Detail)"
        }
        elseif ("$($a.Result)" -eq 'fail' -and "$($b.Result)" -eq 'pass') {
            & $add 'CA gap' 'removed' $b.Title 'resolved'
        }
    }

    # App registrations (by AppId)
    $pApp = _map $Prev.Applications { param($x) $x.AppId }
    $cApp = _map $Curr.Applications { param($x) $x.AppId }
    foreach ($k in $cApp.Keys) { if (-not $pApp.Contains($k)) { & $add 'App registrations' 'added' $cApp[$k].Name '' } }
    foreach ($k in $pApp.Keys) { if (-not $cApp.Contains($k)) { & $add 'App registrations' 'removed' $pApp[$k].Name '' } }

    # Plain return: output enumerates naturally; the caller @()-wraps.
    # (Do NOT comma-wrap here - @(call) would nest the array one level deep.)
    return $L.ToArray()
}

function Get-ChangeLog {
    # Diff every consecutive snapshot pair; newest changes first.
    param($Snapshots)
    $all = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -lt @($Snapshots).Count; $i++) {
        foreach ($c in @(Get-SnapshotChanges $Snapshots[$i - 1] $Snapshots[$i])) { $all.Add($c) }
    }
    $sorted = @($all.ToArray() | Sort-Object { "$($_.Ts)" } -Descending)
    return , $sorted
}

# --------------------------------------------------------------------------- #
# Markdown renderers - deterministic; timestamp appears ONLY in index.md
# --------------------------------------------------------------------------- #

function Write-Docs {
    param($Data, [string]$DocsPath, [int]$WarnDays, $ChangeLog = @(), [int]$SnapCount = 1)

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
    if ($Data.Contains('Anonymized') -and $Data['Anonymized']) {
        $md.Add('> **Anonymized.** People, groups, domains, app names and tenant ids below are pseudonyms, not real values. Counts, dates and settings are real. Conditional Access and Intune **policy names are real** - they are admin-authored text, so check them before sharing this.')
        $md.Add('')
    }
    $md.Add('| Section | At a glance |')
    $md.Add('|---|---|')
    $md.Add("| [1. Tenant](01-tenant.md) | $(@($o.Domains).Count) domain(s), $(@($Data.Licenses).Count) license SKU(s), $($Data.UserCounts.Members) members + $($Data.UserCounts.Guests) guests |")
    $gaps = @(Get-CaGapAnalysis $Data | Where-Object { $_.Result -eq 'fail' })
    $gapsCrit = @($gaps | Where-Object { $_.Severity -eq 'critical' }).Count
    $md.Add("| [2. Conditional Access](02-conditional-access.md) | $caEnabled enabled, $caReport report-only, $caOff disabled; $(@($gaps).Count) gap(s), $gapsCrit critical |")
    $md.Add("| [3. Directory roles](03-roles.md) | $(@($Data.Roles).Count) role(s) with assignments |")
    $md.Add("| [4. Groups](04-groups.md) | $($Data.Groups.Total) total, $(@($Data.Groups.Dynamic).Count) dynamic |")
    $md.Add("| [5. Authentication methods](05-authentication.md) | $(@($Data.AuthMethods | Where-Object { $_.State -eq 'enabled' }).Count) of $(@($Data.AuthMethods).Count) methods enabled |")
    $md.Add("| [6. User & guest settings](06-user-settings.md) | $(@($Data.UserSettings.Keys).Count) settings documented |")
    $md.Add("| [7. Applications](07-applications.md) | $(@($Data.Applications).Count) registrations, $(@($expiring).Count) credential(s) expiring/expired |")
    $intuneGlance = if ($Data.Intune -and $Data.Intune.Available) {
        "$($Data.Intune.Devices.Total) device(s), $(@($Data.Intune.CompliancePolicies).Count) compliance policies, $(@($Data.Intune.ConfigurationProfiles).Count) profiles"
    } else { 'Not available in this snapshot' }
    $md.Add("| [8. Intune](08-intune.md) | $intuneGlance |")
    $md.Add("| [9. Change log](09-changelog.md) | $(@($ChangeLog).Count) change(s) across $SnapCount snapshot(s) |")
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
    $md.Add('')
    $md.Add('## Gap analysis')
    $md.Add('')
    $md.Add('Opinionated baseline hygiene checks, not a compliance audit.')
    $md.Add('')
    $md.Add('| Check | Result | Detail |')
    $md.Add('|---|---|---|')
    foreach ($c in @(Get-CaGapAnalysis $Data)) {
        $r = switch ("$($c.Result)") {
            'pass'    { 'PASS' }
            'fail'    { "GAP ($($c.Severity))" }
            default   { 'UNKNOWN' }
        }
        $md.Add("| $(MdEscape $c.Title) | $r | $(MdEscape $c.Detail) |")
    }
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

    # ---- 08-intune.md ---- #
    $i = $Data.Intune
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 8. Intune')
    $md.Add('')
    if ($i -and $i.Available) {
        $md.Add('## Managed devices')
        $md.Add('')
        $md.Add('| | |')
        $md.Add('|---|---:|')
        foreach ($k in $i.Devices.Keys) { $md.Add("| $k | $($i.Devices[$k]) |") }
        $md.Add('')
        $md.Add('## Compliance state')
        $md.Add('')
        $md.Add('| | |')
        $md.Add('|---|---:|')
        foreach ($k in $i.ComplianceSummary.Keys) { $md.Add("| $k | $($i.ComplianceSummary[$k]) |") }
        $md.Add('')
        $md.Add('## Compliance policies')
        if (@($i.CompliancePolicies).Count) {
            foreach ($p in @($i.CompliancePolicies)) {
                $md.Add('')
                $md.Add("### $(MdEscape $p.Name)")
                $md.Add('')
                $md.Add("Type: $($p.Type) | Assigned to: $(MdEscape ((@($p.Assignments) -join '; ')))")
                if (@($p.Settings.Keys).Count) {
                    $md.Add('')
                    $md.Add('| Setting | Value |')
                    $md.Add('|---|---|')
                    foreach ($k in $p.Settings.Keys) { $md.Add("| $(MdEscape $k) | $(MdEscape $p.Settings[$k]) |") }
                }
            }
        } else { $md.Add(''); $md.Add('None defined.') }
        $md.Add('')
        $md.Add('## Configuration profiles')
        $md.Add('')
        if (@($i.ConfigurationProfiles).Count) {
            $md.Add('| Profile | Type | Assigned to |')
            $md.Add('|---|---|---|')
            foreach ($p in @($i.ConfigurationProfiles)) {
                $md.Add("| $(MdEscape $p.Name) | $($p.Type) | $(MdEscape ((@($p.Assignments) -join '; '))) |")
            }
        } else { $md.Add('None defined.') }
        $md.Add('')
        $md.Add('## App protection policies')
        $md.Add('')
        if (@($i.AppProtection).Count) {
            $md.Add('| Policy | Type |')
            $md.Add('|---|---|')
            foreach ($p in @($i.AppProtection)) { $md.Add("| $(MdEscape $p.Name) | $($p.Type) |") }
        } else { $md.Add('None defined.') }
        $md.Add('')
        $md.Add('Settings catalog policies are not included - that API is still beta-only in Microsoft Graph.')
    } else {
        $md.Add('Not available in this snapshot (no Intune license, missing scopes, or -SkipIntune).')
    }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '08-intune.md') -Encoding UTF8

    # ---- 09-changelog.md ---- #
    # The one section file that carries timestamps by design: it only gains
    # lines when the tenant actually changed between snapshots.
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add('# 9. Change log')
    $md.Add('')
    $md.Add("Computed from $SnapCount archived snapshot(s). Only snapshots where something changed appear below, newest first.")
    if (@($ChangeLog).Count) {
        foreach ($grp in (@($ChangeLog) | Group-Object { $_.Ts } | Sort-Object Name -Descending)) {
            $md.Add('')
            $md.Add("## $($grp.Name)")
            $md.Add('')
            foreach ($c in ($grp.Group | Sort-Object { "$($_.Category)|$($_.Kind)|$($_.Item)" })) {
                $detail = if ($c.Detail) { " - $(MdEscape $c.Detail)" } else { '' }
                $md.Add("- **$($c.Category)** $($c.Kind): $(MdEscape $c.Item)$detail")
            }
        }
    } else {
        $md.Add('')
        $md.Add($(if ($SnapCount -lt 2) { 'History starts here - the change log fills in from the next run onward.' } else { 'No configuration changes detected between snapshots.' }))
    }
    $md -join "`n" | Set-Content -Path (Join-Path $DocsPath '09-changelog.md') -Encoding UTF8
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
header .anon { margin-top: 10px; padding: 9px 12px; border-radius: 8px; font-size: 12.5px;
  line-height: 1.5; color: var(--ink-2); background: var(--surface);
  border: 1px solid var(--border); border-left: 3px solid var(--warning); }
header .anon b { color: var(--ink); font-weight: 650; }
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
.trend-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; }
.trend { background: var(--page); border: 1px solid var(--border); border-radius: 10px; padding: 10px 12px; }
.trend .k { font-size: 12px; color: var(--ink-2); }
.trend .v { font-size: 20px; font-weight: 650; font-variant-numeric: tabular-nums;
  display: flex; align-items: baseline; gap: 8px; }
.trend .d { font-size: 11.5px; color: var(--muted); font-weight: 500; font-variant-numeric: tabular-nums; }
.trend svg { display: block; width: 100%; height: 40px; margin-top: 6px; }
.chg { display: flex; gap: 8px; padding: 5px 0; font-size: 13px; align-items: baseline;
  border-bottom: 1px solid var(--grid); }
.chg:last-child { border-bottom: none; }
.chg .kind { flex: none; width: 76px; font-size: 11.5px; font-weight: 600; color: var(--ink-2);
  text-transform: uppercase; letter-spacing: .3px; }
.chg .cat { color: var(--muted); }
.chg-ts { margin: 12px 0 2px; font-size: 12px; color: var(--muted); font-variant-numeric: tabular-nums; }
[hidden] { display: none !important; }
@media (max-width: 700px) { .bar-row { grid-template-columns: 1fr 1fr 40px; } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1 id="t-name"></h1>
    <div class="sub">Entra ID tenant report &middot; read-only snapshot &middot; <span id="t-id"></span></div>
    <div class="stamp">Generated <b id="t-gen"></b> (UTC)</div>
    <div class="anon" id="t-anon" hidden></div>
  </header>
  <div class="cards" id="kpis"></div>
  <section id="trends-section" hidden><h2>Trends</h2>
    <p class="note" id="trends-note"></p>
    <div class="trend-grid" id="trends"></div>
  </section>
  <section id="changes-section" hidden><h2>What changed</h2>
    <p class="note" id="changes-note"></p>
    <div id="changes"></div>
  </section>
  <section><h2>Conditional Access</h2>
    <p class="note" id="ca-note"></p>
    <div id="ca-gaps" style="margin-bottom:12px"></div>
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
  <section id="intune-section" hidden><h2>Intune</h2>
    <div class="cards" id="intune-tiles" style="margin-bottom:8px"></div>
    <p class="note" id="intune-note"></p>
    <div style="overflow-x:auto"><table id="intune-compliance"></table></div>
    <div id="intune-extra"></div>
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
if (D.Anonymized) {
  const a = document.getElementById('t-anon');
  a.hidden = false;
  a.innerHTML = '<b>Anonymized.</b> People, groups, domains, app names and tenant ids on this page are ' +
    'pseudonyms, not real values. Counts, dates and settings are real. Conditional Access and Intune ' +
    '<b>policy names are real</b> &mdash; they are admin-authored text, so check them before sharing this.';
}

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

// ---- Trends (small multiples; single series each, direct labels, no legend) ----
const series = (D.TrendSeries || []);
if (series.length >= 2) {
  document.getElementById('trends-section').hidden = false;
  document.getElementById('trends-note').textContent = series.length + ' snapshots, ' +
    (series[0].Ts || '').slice(0, 10) + ' \u2192 ' + (series[series.length - 1].Ts || '').slice(0, 10) + '.';
  const METRICS = [
    { k: 'Members', label: 'Members' },
    { k: 'Guests', label: 'Guests' },
    { k: 'CaEnabled', label: 'CA policies enforced' },
    { k: 'RoleAssignments', label: 'Role assignments' },
    { k: 'AppRegistrations', label: 'App registrations' },
    { k: 'CredsInWindow', label: 'Credentials to renew' },
    { k: 'EnrolledDevices', label: 'Enrolled devices' },
    { k: 'NonCompliantDevices', label: 'Non-compliant devices' },
    { k: 'CaGapsCritical', label: 'Critical CA gaps' }
  ];
  function spark(vals) {
    const w = 220, h = 40, pad = 3;
    const min = Math.min(...vals), max = Math.max(...vals);
    const span = max - min;
    const x = i => pad + i * (w - 2 * pad) / Math.max(1, vals.length - 1);
    const y = v => span === 0 ? h / 2 : h - pad - (v - min) * (h - 2 * pad) / span;
    const pts = vals.map((v, i) => x(i).toFixed(1) + ',' + y(v).toFixed(1)).join(' ');
    const area = pad + ',' + (h - pad) + ' ' + pts + ' ' + x(vals.length - 1).toFixed(1) + ',' + (h - pad);
    const lastX = x(vals.length - 1).toFixed(1), lastY = y(vals[vals.length - 1]).toFixed(1);
    return '<svg viewBox="0 0 ' + w + ' ' + h + '" preserveAspectRatio="none" aria-hidden="true">' +
      '<polygon points="' + area + '" fill="var(--accent)" opacity="0.1"></polygon>' +
      '<polyline points="' + pts + '" fill="none" stroke="var(--accent)" stroke-width="2" vector-effect="non-scaling-stroke"></polyline>' +
      '<circle cx="' + lastX + '" cy="' + lastY + '" r="3" fill="var(--accent)" stroke="var(--surface)" stroke-width="2"></circle></svg>';
  }
  document.getElementById('trends').innerHTML = METRICS.map(m => {
    // Trailing run of non-null points: newer sections (e.g. Intune) may not
    // exist in older snapshots - chart only from where the data begins.
    let vals = series.map(r => (r[m.k] === null || r[m.k] === undefined) ? null : Number(r[m.k]));
    let start = vals.length;
    for (let i = vals.length - 1; i >= 0 && vals[i] !== null; i--) start = i;
    vals = vals.slice(start);
    if (vals.length < 2) return '';
    const cur = vals[vals.length - 1], prev = vals[vals.length - 2];
    const d = cur - prev;
    const delta = d === 0 ? 'no change' : (d > 0 ? '+' : '') + d + ' since last';
    return '<div class="trend"><div class="k">' + esc(m.label) + '</div><div class="v">' + esc(cur) +
      ' <span class="d">' + esc(delta) + '</span></div>' + spark(vals) + '</div>';
  }).filter(Boolean).join('');
}

// ---- What changed ----
const changes = (D.Changes || []);
if (changes.length) {
  document.getElementById('changes-section').hidden = false;
  const byTs = {};
  changes.forEach(c => { (byTs[c.Ts] = byTs[c.Ts] || []).push(c); });
  const tss = Object.keys(byTs).sort().reverse();
  document.getElementById('changes-note').textContent = changes.length + ' change(s) across ' +
    tss.length + ' snapshot(s). Full history in docs/09-changelog.md.';
  const KINDWORD = { added: 'Added', removed: 'Removed', changed: 'Changed' };
  const CAP = 4;
  document.getElementById('changes').innerHTML = tss.slice(0, CAP).map(ts =>
    '<div class="chg-ts">' + esc(ts.replace('T', ' ').replace('Z', ' UTC')) + '</div>' +
    byTs[ts].map(c =>
      '<div class="chg"><span class="kind">' + esc(KINDWORD[c.Kind] || c.Kind) + '</span><span>' +
      esc(c.Item) + (c.Detail ? ' \u2014 ' + esc(c.Detail) : '') +
      ' <span class="cat">\u00B7 ' + esc(c.Category) + '</span></span></div>'
    ).join('')
  ).join('') + (tss.length > CAP ? '<p class="note" style="margin-top:8px">Older changes in docs/09-changelog.md.</p>' : '');
} else if (series.length >= 2) {
  document.getElementById('changes-section').hidden = false;
  document.getElementById('changes-note').textContent = 'No configuration changes detected between snapshots.';
}

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
const gaps = (D.CaGaps || []);
if (gaps.length) {
  const SEVCOLOR = { critical: 'critical', warning: 'warning', info: 'muted' };
  const failing = gaps.filter(c => c.Result === 'fail');
  const passing = gaps.filter(c => c.Result !== 'fail');
  const row = c => {
    let b;
    if (c.Result === 'pass') b = badge('good', 'check', 'Pass');
    else if (c.Result === 'unknown') b = '<span class="badge muted">' + ICONS.dash + 'Unknown</span>';
    else if (c.Severity === 'info') b = '<span class="badge muted">' + ICONS.dash + 'Gap</span>';
    else b = badge(SEVCOLOR[c.Severity] || 'warning', 'warn', 'Gap \u00B7 ' + c.Severity);
    return '<div class="chg"><span style="flex:none;width:132px">' + b + '</span><span>' +
      esc(c.Title) + (c.Detail ? ' <span class="muted">\u2014 ' + esc(c.Detail) + '</span>' : '') + '</span></div>';
  };
  document.getElementById('ca-gaps').innerHTML =
    '<div style="font-size:13px;font-weight:600;margin-bottom:2px">Gap analysis <span class="muted" style="font-weight:400">\u00B7 baseline hygiene, not a compliance audit</span></div>' +
    failing.map(row).join('') +
    (passing.length ? '<details><summary>' + passing.length + ' passing / not applicable</summary>' +
      passing.map(row).join('') + '</details>' : '');
}

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

// ---- Intune ----
const IN = D.Intune;
if (IN && IN.Available) {
  document.getElementById('intune-section').hidden = false;
  const dev = IN.Devices || {}, comp = IN.ComplianceSummary || {};
  document.getElementById('intune-tiles').innerHTML = [
    { k: 'Devices', v: dev.Total }, { k: 'Windows', v: dev.Windows },
    { k: 'macOS', v: dev.MacOS }, { k: 'iOS', v: dev.IOS },
    { k: 'Android', v: dev.Android }, { k: 'Non-compliant', v: comp.NonCompliant }
  ].map(x => '<div class="card"><div class="k">' + esc(x.k) + '</div><div class="v">' + esc(x.v == null ? 0 : x.v) +
    '</div></div>').join('');
  const compBits = [badge('good', 'check', (comp.Compliant || 0) + ' compliant')];
  compBits.push((comp.NonCompliant || 0) > 0
    ? badge('critical', 'warn', comp.NonCompliant + ' non-compliant')
    : '<span class="badge muted">' + ICONS.dash + '0 non-compliant</span>');
  if ((comp.InGracePeriod || 0) > 0) compBits.push(badge('warning', 'clock', comp.InGracePeriod + ' in grace period'));
  document.getElementById('intune-note').innerHTML =
    compBits.join(' &nbsp; ') + ' &nbsp; <span class="muted">Full detail in docs/08-intune.md.</span>';
  const pols = (IN.CompliancePolicies || []);
  document.getElementById('intune-compliance').innerHTML = pols.length
    ? '<tr><th>Compliance policy</th><th>Type</th><th>Assigned to</th><th class="num">Settings</th></tr>' +
      pols.map(p => '<tr><td>' + esc(p.Name) + '</td><td>' + esc(p.Type) + '</td><td>' +
        esc((p.Assignments || []).join(', ')) + '</td><td class="num">' +
        Object.keys(p.Settings || {}).length + '</td></tr>').join('')
    : '<tr><td class="muted">No compliance policies defined.</td></tr>';
  let ix = '';
  const profs = (IN.ConfigurationProfiles || []);
  if (profs.length) {
    ix += '<details><summary>Configuration profiles (' + profs.length + ')</summary><ul class="plain">' +
      profs.map(p => '<li>' + esc(p.Name) + ' <span class="muted">\u00B7 ' + esc(p.Type) +
        ' \u00B7 ' + esc((p.Assignments || []).join(', ')) + '</span></li>').join('') + '</ul></details>';
  }
  const mam = (IN.AppProtection || []);
  if (mam.length) {
    ix += '<details><summary>App protection policies (' + mam.length + ')</summary><ul class="plain">' +
      mam.map(p => '<li>' + esc(p.Name) + ' <span class="muted">\u00B7 ' + esc(p.Type) + '</span></li>').join('') +
      '</ul></details>';
  }
  document.getElementById('intune-extra').innerHTML = ix;
}

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
    param($Data, [string]$Path, [int]$WarnDays, $TrendSeries = @(), $ChangeLog = @())
    # The HTML gets the same snapshot plus pre-computed credential rows, so the
    # page needs zero date math and stays consistent with the docs.
    $payload = [ordered]@{}
    foreach ($k in $Data.Keys) { $payload[$k] = $Data[$k] }
    $payload['CredRows'] = @(Get-CredRows $Data $WarnDays | ForEach-Object {
        [ordered]@{ App = $_.App; Type = $_.Type; CredName = $_.CredName
                    ExpiresUtc = $_.ExpiresUtc; DaysLeft = $_.DaysLeft; Severity = $_.Severity }
    })
    $payload['TrendSeries'] = @($TrendSeries)
    $payload['Changes'] = @(@($ChangeLog) | Select-Object -First 200)
    $payload['CaGaps'] = @(Get-CaGapAnalysis $Data)
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
# Anonymization
#
# Applied to the in-memory snapshots AFTER this run's real snapshot has been
# archived and the history has been loaded, and BEFORE anything is rendered.
# That ordering is what makes it safe and coherent: the history on disk stays
# real, and every snapshot in the run shares one mapping table, so the change
# log does not report every admin as removed-and-re-added.
# --------------------------------------------------------------------------- #

$script:AnonSalt   = $null
$script:AnonMap    = @{}   # "kind|real value" -> pseudonym
$script:AnonTaken  = @{}   # pseudonym -> $true, so two real values never merge
$script:AnonDomain = 'example.com'

function Initialize-Anonymizer {
    param([string]$Salt)
    # A fresh salt per run means the output cannot be tested against a guessed
    # name. A fixed salt trades that away for cross-run stability.
    $script:AnonSalt  = if ($Salt) { $Salt } else { [guid]::NewGuid().ToString() }
    $script:AnonMap   = @{}
    $script:AnonTaken = @{}
    $script:AnonDomain = 'example.com'
    $script:AnonDomainLocked = $false
}

function Get-AnonTag {
    param([string]$Kind, [string]$Value, [int]$Chars)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes("$script:AnonSalt|$Kind|$Value")) }
    finally { $sha.Dispose() }
    $hex = -join ($bytes | ForEach-Object { $_.ToString('x2') })
    return $hex.Substring(0, $Chars)
}

function Get-Pseudonym {
    # Template placeholders: {0} = kind, {1} = tag.
    # On collision the tag lengthens rather than silently merging two real
    # identities into one pseudonym - a merge would corrupt the change log.
    param([string]$Kind, $Value, [string]$Template = '{0} {1}')
    if ($null -eq $Value -or "$Value" -eq '') { return $Value }
    $real = "$Value"
    $key  = "$Kind|$real"
    if ($script:AnonMap.ContainsKey($key)) { return $script:AnonMap[$key] }
    for ($chars = 6; $chars -le 32; $chars += 2) {
        $candidate = $Template -f $Kind, (Get-AnonTag $Kind $real $chars)
        if (-not $script:AnonTaken.ContainsKey($candidate)) {
            $script:AnonMap[$key] = $candidate
            $script:AnonTaken[$candidate] = $true
            return $candidate
        }
        if ($script:AnonMap.ContainsKey($key)) { return $script:AnonMap[$key] }
    }
    throw "Anonymizer exhausted the pseudonym space for a '$Kind' value."
}

function Get-AnonGuid {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $Value }
    $t = Get-AnonTag 'Guid' "$Value" 32
    return ($t.Substring(0,8) + '-' + $t.Substring(8,4) + '-' + $t.Substring(12,4) +
            '-' + $t.Substring(16,4) + '-' + $t.Substring(20,12))
}

function Test-IsGuidLike { param($v) return ("$v" -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') }

function Get-AnonDomainName {
    param($Domain)
    if ($null -eq $Domain -or "$Domain" -eq '') { return $Domain }
    # Keep the .onmicrosoft.com shape - the report labels the initial domain,
    # and losing that would make the output read wrong rather than anonymous.
    $suffix = if ("$Domain" -match '\.onmicrosoft\.com$') { '.onmicrosoft.com' } else { '.com' }
    return (Get-Pseudonym 'Domain' $Domain ('example-{1}' + $suffix))
}

function Get-AnonPerson {
    # Display name is the identity key: it is the only thing Conditional Access
    # gives us, so keying on it makes the same human map the same way whether
    # they turn up in a role assignment or a policy exclusion.
    param($DisplayName, $RealUpn, [string]$Kind = 'Person')
    # Note the parameter is $RealUpn, not $Upn: PowerShell variable names are
    # case-insensitive, so a local $upn would silently overwrite the parameter.
    $key = if ($DisplayName -and "$DisplayName" -ne '') { "$DisplayName" } else { "$RealUpn" }
    if (-not $key) { return @{ Name = $DisplayName; Upn = $RealUpn } }
    $name = Get-Pseudonym $Kind $key
    $newUpn = $null
    if ($RealUpn -and "$RealUpn" -ne '') {
        $tag = ("$name" -split ' ')[-1]
        $newUpn = "$($Kind.ToLowerInvariant())-$tag@$script:AnonDomain"
    }
    return @{ Name = $name; Upn = $newUpn }
}

$script:ANON_KEEP_PRINCIPALS = @('All', 'None', 'GuestsOrExternalUsers')
$script:ANON_KEEP_APPS = @('All cloud apps', 'None', 'Office 365', 'Microsoft Admin Portals',
                           'All', 'Office365', 'MicrosoftAdminPortals')
$script:ANON_KEEP_TARGETS = @('all_users', 'all_devices')

function Get-AnonPrincipalList {
    param($Values, [string]$Kind)
    return @(foreach ($v in @($Values)) {
        if ($null -eq $v -or "$v" -eq '') { continue }
        elseif ("$v" -in $script:ANON_KEEP_PRINCIPALS) { "$v" }
        elseif (Test-IsGuidLike $v) { Get-AnonGuid $v }   # stayed unresolved - keep it honest
        else { Get-Pseudonym $Kind $v }
    })
}

function Get-AnonScrubbedString {
    # Targeted scrub for free-text values we otherwise keep: anything that
    # looks like an address is replaced, everything else survives.
    param($Value)
    if ($null -eq $Value) { return $Value }
    $s = "$Value"
    $s = [regex]::Replace($s, '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', { param($m) "person-$(Get-AnonTag 'Scrub' $m.Value 6)@$script:AnonDomain" })
    $s = [regex]::Replace($s, 'https?://[^\s"'')]+', { param($m) "https://example-$(Get-AnonTag 'Scrub' $m.Value 6).com" })
    $s = [regex]::Replace($s, '\\\\[^\s"'')]+', { param($m) "\\server-$(Get-AnonTag 'Scrub' $m.Value 6)\share" })
    $s = [regex]::Replace($s, '\b\d{1,3}(\.\d{1,3}){3}\b', { param($m) '203.0.113.0' })
    return $s
}

function Get-AnonMembershipRule {
    # Keep the rule's shape - attributes and operators are Microsoft vocabulary
    # and the interesting part - and replace only the quoted literals, which is
    # where department names, domains and cost centres live.
    param($Rule)
    if ($null -eq $Rule -or "$Rule" -eq '') { return $Rule }
    return [regex]::Replace("$Rule", '"([^"]*)"', {
        param($m)
        if ($m.Groups[1].Value -eq '') { '""' }
        else { '"' + (Get-Pseudonym 'Value' $m.Groups[1].Value 'value-{1}') + '"' }
    })
}

function Get-AnonIntuneAssignments {
    param($Assignments)
    return @(foreach ($a in @($Assignments)) {
        $s = "$a"
        if ($s -eq '') { continue }
        elseif ($s -in @('All devices', 'All users')) { $s }
        elseif ($s -like 'Exclude: *') {
            $inner = $s.Substring(9)
            if (Test-IsGuidLike $inner) { "Exclude: $(Get-AnonGuid $inner)" }
            else { "Exclude: $(Get-Pseudonym 'Group' $inner)" }
        }
        elseif (Test-IsGuidLike $s) { Get-AnonGuid $s }
        elseif ($s -like '*AssignmentTarget') { $s }   # an odata type we could not name
        else { Get-Pseudonym 'Group' $s }
    })
}

function Protect-TenantData {
    <# In-place. Idempotent guard: a snapshot already marked Anonymized is left
       alone, so passing the same object twice cannot double-map it. #>
    param($Data)
    if (-not $Data) { return }
    if ($Data.Contains('Anonymized') -and $Data['Anonymized']) { return }

    # -- domains first: the UPN suffix depends on them --------------------- #
    $org = $Data['Organization']
    if ($org) {
        $mapped = @()
        foreach ($d in @($org['Domains'])) {
            $d['Name'] = Get-AnonDomainName $d['Name']
            $mapped += , $d
        }
        $org['Domains'] = @($mapped)
        # Locked to the first snapshot processed (the current one), so a domain
        # that existed only in old history cannot shift the UPN suffix midway
        # and split one person across two spellings.
        if (-not $script:AnonDomainLocked) {
            $default = @($mapped | Where-Object { $_['IsDefault'] })
            $script:AnonDomain = if ($default.Count) { "$($default[0]['Name'])" }
                                 elseif ($mapped.Count) { "$($mapped[0]['Name'])" }
                                 else { 'example.com' }
            $script:AnonDomainLocked = $true
        }
        $org['DisplayName'] = Get-Pseudonym 'Org' $org['DisplayName'] 'Example Organization {1}'
    }
    $Data['TenantId'] = Get-AnonGuid $Data['TenantId']

    # -- Conditional Access ------------------------------------------------- #
    $ca = $Data['ConditionalAccess']
    if ($ca) {
        foreach ($p in @($ca['Policies'])) {
            # Policy Name is deliberately left as written - see -Anonymize help.
            $p['IncludeUsers']  = @(Get-AnonPrincipalList $p['IncludeUsers']  'Person')
            $p['ExcludeUsers']  = @(Get-AnonPrincipalList $p['ExcludeUsers']  'Person')
            $p['IncludeGroups'] = @(Get-AnonPrincipalList $p['IncludeGroups'] 'Group')
            $p['ExcludeGroups'] = @(Get-AnonPrincipalList $p['ExcludeGroups'] 'Group')
            $p['IncludeApps']   = @(foreach ($a in @($p['IncludeApps'])) {
                if ("$a" -in $script:ANON_KEEP_APPS) { "$a" }
                elseif (Test-IsGuidLike $a) { Get-AnonGuid $a } else { Get-Pseudonym 'App' $a } })
            $p['ExcludeApps']   = @(foreach ($a in @($p['ExcludeApps'])) {
                if ("$a" -in $script:ANON_KEEP_APPS) { "$a" }
                elseif (Test-IsGuidLike $a) { Get-AnonGuid $a } else { Get-Pseudonym 'App' $a } })
            if ($p['Locations']) {
                $p['Locations']['Include'] = @(foreach ($l in @($p['Locations']['Include'])) {
                    if ("$l" -in @('All', 'AllTrusted')) { "$l" }
                    elseif (Test-IsGuidLike $l) { Get-AnonGuid $l } else { Get-Pseudonym 'Location' $l } })
                $p['Locations']['Exclude'] = @(foreach ($l in @($p['Locations']['Exclude'])) {
                    if ("$l" -in @('All', 'AllTrusted')) { "$l" }
                    elseif (Test-IsGuidLike $l) { Get-AnonGuid $l } else { Get-Pseudonym 'Location' $l } })
            }
        }
        foreach ($l in @($ca['NamedLocations'])) {
            $l['Name'] = Get-Pseudonym 'Location' $l['Name']
            # Detail is "N range(s), trusted" or a country list - no addresses.
        }
    }

    # -- Directory roles ---------------------------------------------------- #
    foreach ($r in @($Data['Roles'])) {
        foreach ($m in @($r['Members'])) {
            # A role member can be a user, a group, or a service principal -
            # label each as what it is, and share the group map so the same
            # group reads the same here and in a CA policy.
            $kind = switch ("$($m['Type'])") {
                'servicePrincipal' { 'Service' }
                'group'            { 'Group' }
                default            { 'Person' }
            }
            $pair = Get-AnonPerson $m['DisplayName'] $m['UserPrincipalName'] $kind
            $m['DisplayName'] = $pair.Name
            $m['UserPrincipalName'] = $pair.Upn
        }
    }

    # -- Groups -------------------------------------------------------------- #
    $g = $Data['Groups']
    if ($g) {
        $g['RoleAssignable'] = @(foreach ($n in @($g['RoleAssignable'])) { Get-Pseudonym 'Group' $n })
        foreach ($d in @($g['Dynamic'])) {
            $d['Name'] = Get-Pseudonym 'Group' $d['Name']
            $d['Rule'] = Get-AnonMembershipRule $d['Rule']
        }
    }

    # -- Authentication methods ---------------------------------------------- #
    foreach ($m in @($Data['AuthMethods'])) {
        $m['Targets'] = @(foreach ($t in @($m['Targets'])) {
            if ("$t" -in $script:ANON_KEEP_TARGETS) { "$t" }
            elseif (Test-IsGuidLike $t) { Get-AnonGuid $t } else { Get-Pseudonym 'Group' $t } })
    }

    # -- App registrations ---------------------------------------------------- #
    foreach ($a in @($Data['Applications'])) {
        $a['Name']  = Get-Pseudonym 'App' $a['Name']
        $a['AppId'] = Get-AnonGuid $a['AppId']
        foreach ($c in @($a['Credentials'])) {
            if ($c['Name']) { $c['Name'] = Get-Pseudonym 'Credential' $c['Name'] }
        }
    }

    # -- Intune ---------------------------------------------------------------- #
    $i = $Data['Intune']
    if ($i -and $i['Available']) {
        foreach ($p in @($i['CompliancePolicies'])) {
            # Policy Name kept, as above.
            $p['Assignments'] = @(Get-AnonIntuneAssignments $p['Assignments'])
            if ($p['Settings']) {
                foreach ($k in @($p['Settings'].Keys)) {
                    if ($p['Settings'][$k] -is [string]) {
                        $p['Settings'][$k] = Get-AnonScrubbedString $p['Settings'][$k]
                    }
                }
            }
        }
        foreach ($p in @($i['ConfigurationProfiles'])) {
            $p['Assignments'] = @(Get-AnonIntuneAssignments $p['Assignments'])
        }
    }

    $Data['Anonymized'] = $true
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
    Get-TenantData -StaleCredDays $StaleCredDays -SkipIntune:$SkipIntune
}

Normalize-DataDates $data
$null = New-Item -ItemType Directory -Path $OutputPath -Force

$historyDir = if ($HistoryPath) { $HistoryPath }
              elseif ($SampleData) { Join-Path $PSScriptRoot 'sample-history' }
              else { Join-Path $OutputPath 'history' }
if (-not $NoHistory -and -not $FromJson) {
    $savedTo = Save-Snapshot $data $historyDir
    Write-Host "Snapshot archived: $savedTo"
}
$snaps = Get-History $historyDir $data

if ($Anonymize) {
    # After archiving (history on disk stays real), before anything is derived
    # or rendered. Every snapshot in this run shares one mapping table, so the
    # change log and trends stay coherent across the whole series.
    Initialize-Anonymizer -Salt $AnonymizeSalt
    Protect-TenantData $data
    foreach ($s in @($snaps)) { Protect-TenantData $s }
    Write-Host 'Anonymized: identities, groups, domains and app names replaced with pseudonyms.'
    Write-Host '  Policy names are KEPT so the gap analysis stays readable - read them before sharing.'
}

$trend = Get-TrendSeries $snaps $StaleCredDays
$changeLog = Get-ChangeLog $snaps

Write-Docs   -Data $data -DocsPath (Join-Path $OutputPath 'docs') -WarnDays $StaleCredDays `
             -ChangeLog $changeLog -SnapCount @($snaps).Count
$data | ConvertTo-Json -Depth 12 | Set-Content -Path (Join-Path $OutputPath 'tenant.json') -Encoding UTF8
Write-Report -Data $data -Path (Join-Path $OutputPath 'report.html') -WarnDays $StaleCredDays `
             -TrendSeries $trend -ChangeLog $changeLog

# run-summary.json: this run's delta + headline numbers, purpose-built for
# automation (see Send-TenantDocsAlert.ps1) and anything else consuming runs.
$credRowsNow = @(Get-CredRows $data $StaleCredDays)
$latest = @($trend)[-1]
$newChanges = @(@($changeLog) | Where-Object { "$($_.Ts)" -eq "$($data.GeneratedUtc)" })
[ordered]@{
    GeneratedUtc  = $data.GeneratedUtc
    Tenant        = [ordered]@{ Name = $data.Organization.DisplayName; Id = $data.TenantId }
    SnapshotCount = @($snaps).Count
    Kpis          = [ordered]@{
        Members          = $latest.Members
        EnabledMembers   = $latest.EnabledMembers
        Guests           = $latest.Guests
        CaEnabled        = $latest.CaEnabled
        CaTotal          = $latest.CaTotal
        RoleAssignments  = $latest.RoleAssignments
        Groups           = $latest.Groups
        AppRegistrations = $latest.AppRegistrations
        CredsInWindow    = $latest.CredsInWindow
        CredsExpired     = @($credRowsNow | Where-Object { $_.Severity -eq 'expired' }).Count
        CaGapsCritical   = $latest.CaGapsCritical
        CaGapsWarning    = @(Get-CaGapAnalysis $data | Where-Object { $_.Result -eq 'fail' -and $_.Severity -eq 'warning' }).Count
        EnrolledDevices  = $latest.EnrolledDevices
        NonCompliantDevices = $latest.NonCompliantDevices
    }
    NewChanges    = $newChanges
    CaGaps        = @(Get-CaGapAnalysis $data)
} | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $OutputPath 'run-summary.json') -Encoding UTF8

Write-Host ''
Write-Host "Done. Output in $OutputPath`:"
Write-Host '  docs/         10 Markdown files (commit these - the git diff is your drift report)'
Write-Host '  tenant.json   the full snapshot'
Write-Host '  report.html   the shareable report - open it in a browser'
Write-Host ("  history       {0} snapshot(s) -> trends + change log" -f @($snaps).Count)
Write-Host ("  run-summary   {0} new change(s) this run - feed it to Send-TenantDocsAlert.ps1" -f @($newChanges).Count)

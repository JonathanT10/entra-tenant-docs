# Test: a section the sign-in cannot reach is SKIPPED, not attempted.
#     pwsh tests/test-intune-scope.ps1
#
# This exists because of a real incident. A refresh printed
#
#     Intune unavailable (InteractiveBrowserCredential authentication failed:
#     User canceled authentication. ) - section skipped
#
# which means the collector called an Intune endpoint the sign-in was never
# granted, the Graph SDK went off to acquire the scope on the spot, and a
# browser window opened in the middle of the run - behind the window the
# person was actually watching. Nobody sees it, nobody finishes it, and the
# refresh sits there looking hung.
#
# Nothing here talks to Microsoft 365: Graph is stubbed, and every URI the
# collector asks for is logged so "never attempted" can be proven rather than
# asserted.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath
$repo = Split-Path -Parent $here
$work = Join-Path ([IO.Path]::GetTempPath()) "intune-scope-$([guid]::NewGuid().ToString('n').Substring(0,6))"
$null = New-Item -ItemType Directory -Path $work -Force

$fails = [System.Collections.Generic.List[string]]::new()
function Check { param([string]$Label, [bool]$Cond)
    Write-Host ("{0} {1}" -f ($(if ($Cond) { 'PASS' } else { 'FAIL' })), $Label)
    if (-not $Cond) { $fails.Add($Label) }
}

$preamble = @'
function Get-MgContext {
    [pscustomobject]@{ Scopes = @($env:TD_SCOPES -split ','); TenantId = 'test-tenant'; Account = 'tester@example.com' }
}
function Connect-MgGraph { Add-Content -Path $env:TD_CALLLOG -Value 'CONNECT-ATTEMPTED' }
function Invoke-MgGraphRequest {
    param([string]$Method, [string]$Uri, [string]$OutputType, $Body)
    Add-Content -Path $env:TD_CALLLOG -Value $Uri
    # An Intune call must never get here. If one does, behave the way the live
    # SDK did - fail the way a cancelled sign-in window fails - so the test
    # sees the same symptom the incident produced.
    if ($Uri -like '*deviceManagement*' -or $Uri -like '*deviceAppManagement*') {
        throw 'InteractiveBrowserCredential authentication failed: User canceled authentication. '
    }
    # Enough shape for the collection to get as far as the Intune section.
    # Every list call gets one plausible row; the organization call gets the
    # handful of fields the renderer indexes into.
    $row = [pscustomobject]@{
        id = '00000000-0000-0000-0000-000000000001'
        displayName = 'Thing'; userPrincipalName = 'u@example.com'
        accountEnabled = $true; userType = 'Member'
        createdDateTime = (Get-Date).AddDays(-30).ToString('o')
        state = 'enabled'; appId = '11111111-1111-1111-1111-111111111111'
        keyCredentials = @(); passwordCredentials = @()
        conditions = $null; grantControls = $null; sessionControls = $null
        verifiedDomains = @([pscustomobject]@{ name = 'example.com'; isDefault = $true; isInitial = $true })
        assignedPlans = @(); onPremisesSyncEnabled = $false
        roleTemplateId = $null; members = @()
        groupTypes = @(); membershipRule = $null; isAssignableToRole = $false
        securityEnabled = $true; mailEnabled = $false
        skuPartNumber = 'TEST'; prepaidUnits = [pscustomobject]@{ enabled = 1 }; consumedUnits = 0
        servicePlans = @()
        roleDefinitionId = '22222222-2222-2222-2222-222222222222'
        principal = [pscustomobject]@{
            '@odata.type' = '#microsoft.graph.user'
            id = '33333333-3333-3333-3333-333333333333'
            displayName = 'Person'; userPrincipalName = 'person@example.com'
        }
        ipRanges = @(); countriesAndRegions = @()
        signInActivity = $null
    }
    [pscustomobject]@{ value = @($row); '@odata.nextLink' = $null }
}
'@

function Run-Collect {
    param([string]$Scopes)
    $log = Join-Path $work "calls-$([guid]::NewGuid().ToString('n').Substring(0,4)).txt"
    Set-Content -Path $log -Value ''
    $env:TD_SCOPES = $Scopes; $env:TD_CALLLOG = $log
    $out = Join-Path $work "out-$([guid]::NewGuid().ToString('n').Substring(0,4))"
    $cmd = $preamble + "`n& '$repo/Export-EntraTenantDocs.ps1' -OutputPath '$out' -NoHistory"
    $text = (& pwsh -NoProfile -Command $cmd 2>&1 | Out-String)
    return @{ Text = $text; Calls = @(Get-Content $log | Where-Object { $_ }) }
}

$BASE = 'Directory.Read.All,Policy.Read.All,RoleManagement.Read.Directory,Application.Read.All,Organization.Read.All,User.Read.All'
$INTUNE = 'DeviceManagementConfiguration.Read.All,DeviceManagementManagedDevices.Read.All,DeviceManagementApps.Read.All'

Write-Host ''
Write-Host '-- 1. Intune scopes NOT granted: the endpoint is never called'
$r = Run-Collect -Scopes $BASE
Check 'no Intune endpoint was called' (@($r.Calls | Where-Object { $_ -like '*deviceManagement*' -or $_ -like '*deviceAppManagement*' }).Count -eq 0)
Check 'it says why, in plain words' ($r.Text -like '*not granted the Intune permissions*')
Check 'no browser sign-in was provoked' ($r.Text -notlike '*InteractiveBrowserCredential*' -and $r.Text -notlike '*User canceled*')
Check 'it did not try to re-connect behind your back' (@($r.Calls | Where-Object { $_ -eq 'CONNECT-ATTEMPTED' }).Count -eq 0)
Check 'the rest of the collection still ran' (@($r.Calls | Where-Object { $_ -like '*organization*' }).Count -gt 0)

Write-Host ''
Write-Host '-- 2. Intune scopes granted: the endpoint IS called'
$r = Run-Collect -Scopes "$BASE,$INTUNE"
Check 'the Intune endpoint is reached when the sign-in allows it' (@($r.Calls | Where-Object { $_ -like '*deviceManagement*' }).Count -gt 0)
Check 'and a genuine failure there is still reported, not hidden' ($r.Text -like '*Intune unavailable*')

Write-Host ''
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
if ($fails.Count) { Write-Host "RESULT: $($fails.Count) FAILURES"; $fails | ForEach-Object { Write-Host "  - $_" }; exit 1 }
Write-Host 'RESULT: ALL PASS'
exit 0

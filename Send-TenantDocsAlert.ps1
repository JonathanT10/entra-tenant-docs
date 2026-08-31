<#
.SYNOPSIS
    Turn a tenant-docs run into an alert - Teams and/or email - but only
    when something actually changed.

.DESCRIPTION
    Reads the run-summary.json that Export-EntraTenantDocs.ps1 writes and
    decides whether the run is worth an interruption:

      - the tenant changed (the run's change log gained entries), or
      - an app credential is now EXPIRED, or
      - you passed -AlwaysNotify (daily-digest style)

    If none of those hold, it prints "nothing to report" and exits quietly -
    schedule it as often as you like without training people to ignore it.

    Notification channels (use either or both):

      Teams  - posts an Adaptive Card to a Power Automate Workflows webhook.
               In Teams: channel > Workflows > "Post to a channel when a
               webhook request is received", then copy the URL. (Legacy
               Office 365 connector webhooks were retired in May 2026 -
               this script targets the Workflows format only.)
      Email  - plain-text mail via your internal SMTP relay (unauthenticated;
               point it at a relay that allows the sending host).

    Keep the webhook URL out of scripts and task definitions: set it once in
    the TENANTDOCS_TEAMS_WEBHOOK environment variable (machine scope for
    scheduled tasks) and omit -TeamsWebhookUrl entirely.

    This is the only component in the toolkit that sends anything anywhere -
    and it only talks to YOUR webhook and YOUR relay. The tenant itself is
    never touched.

.PARAMETER RunSummaryPath
    Path to the run-summary.json produced by the export.
    Default: .\tenant-docs\run-summary.json

.PARAMETER TeamsWebhookUrl
    Power Automate Workflows webhook URL. Default: the
    TENANTDOCS_TEAMS_WEBHOOK environment variable.

.PARAMETER SmtpServer
    Internal SMTP relay host. Email is skipped when omitted.

.PARAMETER SmtpPort
    Relay port. Default: 25.

.PARAMETER From
    Sender address for email alerts.

.PARAMETER To
    One or more recipient addresses.

.PARAMETER UseSsl
    Use SSL/TLS to the relay.

.PARAMETER ReportLink
    Optional link shown in the alert - an intranet URL or UNC path where
    report.html lives.

.PARAMETER AlwaysNotify
    Send even when nothing changed (turns the alert into a heartbeat/digest).

.PARAMETER MaxItems
    Cap on change lines in the alert body; the rest are summarized.
    Default: 20.

.EXAMPLE
    .\Export-EntraTenantDocs.ps1 -OutputPath C:\tenant-docs
    .\Send-TenantDocsAlert.ps1 -RunSummaryPath C:\tenant-docs\run-summary.json

    The scheduled-task pair: document, then alert only if something changed.
    (Webhook URL comes from TENANTDOCS_TEAMS_WEBHOOK.)

.EXAMPLE
    .\Send-TenantDocsAlert.ps1 -SmtpServer relay.corp.local -From tenantdocs@corp.com -To itops@corp.com -AlwaysNotify

    Email-only daily digest through an internal relay.
#>
[CmdletBinding()]
param(
    [string]$RunSummaryPath = '.\tenant-docs\run-summary.json',
    [string]$TeamsWebhookUrl = $env:TENANTDOCS_TEAMS_WEBHOOK,
    [string]$SmtpServer,
    [ValidateRange(1, 65535)][int]$SmtpPort = 25,
    [string]$From,
    [string[]]$To,
    [switch]$UseSsl,
    [string]$ReportLink,
    [switch]$AlwaysNotify,
    [ValidateRange(1, 200)][int]$MaxItems = 20
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $RunSummaryPath)) {
    throw "Run summary not found: $RunSummaryPath (run Export-EntraTenantDocs.ps1 first)"
}
$summary = Get-Content $RunSummaryPath -Raw | ConvertFrom-Json

# ConvertFrom-Json parses the ISO timestamp into [datetime]; render it back
# canonically so alerts show UTC ISO, not the runner's locale format.
$genUtc = if ($summary.GeneratedUtc -is [datetime]) {
    $summary.GeneratedUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
} else { "$($summary.GeneratedUtc)" }

$tenantName = "$($summary.Tenant.Name)"
$changes = @($summary.NewChanges)
$expired = [int]$summary.Kpis.CredsExpired
$inWindow = [int]$summary.Kpis.CredsInWindow

# ---- is this worth an interruption? ---------------------------------------- #
$reasons = @()
if ($changes.Count) { $reasons += "$($changes.Count) configuration change(s)" }
if ($expired -gt 0) { $reasons += "$expired EXPIRED app credential(s)" }
if (-not $reasons -and -not $AlwaysNotify) {
    Write-Host "Nothing to report for $tenantName ($genUtc) - no alert sent."
    return [pscustomobject]@{ Notified = $false; Reasons = @(); Teams = 'skipped'; Email = 'skipped' }
}
if (-not $reasons) { $reasons = @('scheduled digest (-AlwaysNotify)') }

# ---- compose the message ---------------------------------------------------- #
$title = if ($changes.Count) { "Tenant change alert - $tenantName" }
         elseif ($expired -gt 0) { "Credential alert - $tenantName" }
         else { "Tenant docs digest - $tenantName" }

$lines = New-Object System.Collections.Generic.List[string]
foreach ($c in ($changes | Select-Object -First $MaxItems)) {
    $detail = if ($c.Detail) { " - $($c.Detail)" } else { '' }
    $lines.Add(("[{0}] {1}: {2}{3}" -f $c.Kind.ToUpper(), $c.Category, $c.Item, $detail))
}
if ($changes.Count -gt $MaxItems) {
    $lines.Add("(+$($changes.Count - $MaxItems) more - see docs/08-changelog.md)")
}
if ($expired -gt 0)  { $lines.Add("$expired app credential(s) EXPIRED - see the report's App credentials section.") }
elseif ($inWindow -gt 0) { $lines.Add("$inWindow app credential(s) inside the renewal window.") }
$lines.Add(("Snapshot {0} | members {1} | guests {2} | CA enforced {3}/{4}" -f `
    $genUtc, $summary.Kpis.Members, $summary.Kpis.Guests, $summary.Kpis.CaEnabled, $summary.Kpis.CaTotal))

$textBody = ($lines.ToArray() -join "`n")
if ($ReportLink) { $textBody += "`nReport: $ReportLink" }

# ---- Teams (Power Automate Workflows webhook, Adaptive Card) ---------------- #
$teamsStatus = 'skipped'
if ($TeamsWebhookUrl) {
    $body = New-Object System.Collections.Generic.List[object]
    $body.Add(@{ type = 'TextBlock'; text = $title; weight = 'Bolder'; size = 'Medium'; wrap = $true })
    $body.Add(@{ type = 'TextBlock'; text = ($reasons -join '; '); isSubtle = $true; wrap = $true; spacing = 'None' })
    foreach ($l in $lines) { $body.Add(@{ type = 'TextBlock'; text = $l; wrap = $true; spacing = 'Small' }) }
    $card = [ordered]@{
        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
        type      = 'AdaptiveCard'
        version   = '1.4'
        body      = $body.ToArray()
    }
    if ($ReportLink -and $ReportLink -match '^https?://') {
        $card['actions'] = @(@{ type = 'Action.OpenUrl'; title = 'Open the report'; url = $ReportLink })
    }
    $payload = @{
        type        = 'message'
        attachments = @(@{ contentType = 'application/vnd.microsoft.card.adaptive'; content = $card })
    } | ConvertTo-Json -Depth 12
    try {
        $null = Invoke-RestMethod -Method Post -Uri $TeamsWebhookUrl -ContentType 'application/json' -Body $payload
        $teamsStatus = 'sent'
        Write-Host "Teams alert sent ($($reasons -join '; '))."
    }
    catch {
        $teamsStatus = "failed: $($_.Exception.Message)"
        Write-Warning "Teams webhook post failed: $($_.Exception.Message)"
    }
}

# ---- Email (internal relay) -------------------------------------------------- #
$emailStatus = 'skipped'
if ($SmtpServer) {
    if (-not $From -or -not $To) {
        Write-Warning 'Email skipped: -SmtpServer needs -From and -To.'
        $emailStatus = 'skipped: missing From/To'
    }
    else {
        try {
            $mailArgs = @{
                SmtpServer = $SmtpServer; Port = $SmtpPort
                From = $From; To = $To
                Subject = "$title ($($reasons -join '; '))"
                Body = $textBody
            }
            if ($UseSsl) { $mailArgs.UseSsl = $true }
            Send-MailMessage @mailArgs -WarningAction SilentlyContinue
            $emailStatus = 'sent'
            Write-Host "Email alert sent to $($To -join ', ')."
        }
        catch {
            $emailStatus = "failed: $($_.Exception.Message)"
            Write-Warning "Email send failed: $($_.Exception.Message)"
        }
    }
}

if ($teamsStatus -eq 'skipped' -and $emailStatus -eq 'skipped') {
    Write-Warning 'No channel configured - pass -TeamsWebhookUrl (or set TENANTDOCS_TEAMS_WEBHOOK) and/or -SmtpServer.'
}

[pscustomobject]@{
    Notified = ($teamsStatus -eq 'sent' -or $emailStatus -eq 'sent')
    Reasons  = $reasons
    Teams    = $teamsStatus
    Email    = $emailStatus
}

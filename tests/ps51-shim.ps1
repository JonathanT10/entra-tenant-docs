# Simulates Windows PowerShell 5.1's ConvertFrom-Json - no -AsHashtable, output
# is a PSCustomObject graph - so the fallback converter is exercised under
# pwsh. A global FUNCTION wins command resolution over the cmdlet, and
# Get-Command returns this shim, which lacks the parameter, so the script's
# feature detection takes the 5.1 path exactly as it would on a real 5.1 box.
function global:ConvertFrom-Json {
    param([Parameter(ValueFromPipeline = $true)][string]$InputObject)
    begin { $chunks = New-Object System.Collections.Generic.List[string] }
    process { if ($null -ne $InputObject) { $chunks.Add($InputObject) } }
    end { Microsoft.PowerShell.Utility\ConvertFrom-Json ($chunks -join "`n") }
}

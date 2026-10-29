<#
.SYNOPSIS
  Merge latest assessment JSON files into a scored posture summary.
#>
param([switch]$Live)
. "$PSScriptRoot\00_Config.ps1"

# Ensure assessment artefacts exist (synthetic default)
& "$PSScriptRoot\01_Assess_Identity_MFA.ps1" @PSBoundParameters
& "$PSScriptRoot\02_Assess_Admins_Guests.ps1" @PSBoundParameters
& "$PSScriptRoot\03_Assess_Audit_Forwarding_SignIn.ps1" @PSBoundParameters

$latest = Get-ChildItem $EvidenceDir -Filter "*_Assessment.json" |
    Sort-Object LastWriteTime -Descending |
    Group-Object { ($_.Name -replace '^\d{8}_\d{6}_', '') } |
    ForEach-Object { $_.Group | Select-Object -First 1 }

$allItems = @()
foreach ($f in $latest) {
    $obj = Get-Content $f.FullName -Raw | ConvertFrom-Json
    foreach ($item in $obj.Items) {
        $allItems += [pscustomobject]@{
            Family   = $obj.ControlFamily
            Mode     = $obj.Mode
            Id       = $item.Id
            Title    = $item.Title
            Severity = $item.Severity
            Score    = [int]$item.Score
            MaxScore = [int]$item.MaxScore
            Observed = $item.Observed
        }
    }
}

$totalScore = ($allItems | Measure-Object -Property Score -Sum).Sum
$totalMax   = ($allItems | Measure-Object -Property MaxScore -Sum).Sum
$pct = if ($totalMax -gt 0) { [math]::Round(100.0 * $totalScore / $totalMax, 1) } else { 0 }

function Get-RiskBand([double]$p) {
    if ($p -ge 80) { return "Low residual (monitor)" }
    if ($p -ge 55) { return "Moderate — prioritise High/Critical" }
    if ($p -ge 30) { return "Elevated — remediation plan required" }
    return "Critical — immediate governance attention"
}

$summary = [ordered]@{
    Organisation     = $FictionalOrg
    GeneratedUtc     = (Get-Date).ToUniversalTime().ToString("o")
    Mode             = (($allItems | Select-Object -ExpandProperty Mode -Unique) -join ",")
    ControlCount     = $allItems.Count
    AggregateScore   = $totalScore
    AggregateMax     = $totalMax
    PosturePercent   = $pct
    RiskBand         = Get-RiskBand $pct
    SeverityCounts   = @{
        Critical = @($allItems | Where-Object Severity -eq "Critical").Count
        High     = @($allItems | Where-Object Severity -eq "High").Count
        Medium   = @($allItems | Where-Object Severity -eq "Medium").Count
        Low      = @($allItems | Where-Object Severity -eq "Low").Count
    }
    Findings         = $allItems
    RemediationTop5  = @(
        "Remove disabled accounts from privileged groups; verify leavers weekly.",
        "Enforce MFA for all admins; prefer phishing-resistant methods where licensed.",
        "Reduce standing Global Admins; use break-glass + time-boxed elevation (see JIT pilot).",
        "Scan for external forwarding and inbox rules; alert on new rules.",
        "Document Free-tier limits: no CA / ID Protection — do not claim those controls."
    )
}

Write-EvidenceJson -NamePrefix "Posture_Summary" -Object $summary | Out-Null
Write-Host ("Posture {0}% — {1}" -f $pct, $summary.RiskBand)

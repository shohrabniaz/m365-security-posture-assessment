<#
.SYNOPSIS
  Assess audit logging, mailbox forwarding, and risky sign-in related controls.
#>
param([switch]$Live)
. "$PSScriptRoot\00_Config.ps1"

$useLive = Assert-LabTenantOrSynthetic -Live:$Live
$findings = [ordered]@{
    ControlFamily = "Audit logging, mailbox forwarding, risky sign-in"
    Mode          = if ($useLive) { "LiveGraph" } else { "Synthetic" }
    Organisation  = $FictionalOrg
    CollectedUtc  = (Get-Date).ToUniversalTime().ToString("o")
    GuidanceRefs  = @(
        "Microsoft: Turn on mailbox auditing / unified audit log",
        "Microsoft: Detect and remediate external mailbox forwarding",
        "Microsoft Entra ID Protection / risky sign-ins (licence-dependent)",
        "ASD: Logging and monitoring guidance for identity"
    )
}

if ($useLive) {
    $findings.Note = "Live path needs Exchange Online / audit APIs and appropriate roles. Identity Protection risky-sign-in APIs need P2 — Free tier must report NotAvailable not Pass."
    $findings.Items = @()
}
else {
    $findings.Items = @(
        [ordered]@{
            Id = "AUD-01"; Title = "Unified audit log enabled"
            Observed = "Audit logging assumed enabled (default in modern tenants) but retention and review process not evidenced"
            Severity = "Medium"; Score = 55; MaxScore = 100
            EvidenceNote = "Synthetic — process gap, not a claimed portal screenshot"
        },
        [ordered]@{
            Id = "FWD-01"; Title = "External mailbox forwarding"
            Observed = "No automated weekly scan of inbox rules / forwarding SMTP in place"
            Severity = "High"; Score = 25; MaxScore = 100
            EvidenceNote = "Synthetic — recommend Transport Rule alert + monthly export"
        },
        [ordered]@{
            Id = "RSK-01"; Title = "Risky sign-in detection"
            Observed = "Entra ID Protection / risk-based CA NotAvailable on Free tier"
            Severity = "High"; Score = 0; MaxScore = 100
            EvidenceNote = "Honest Free-tier limit — do not claim Identity Protection"
        },
        [ordered]@{
            Id = "RSK-02"; Title = "Sign-in log review cadence"
            Observed = "No documented weekly review of failed/unusual sign-ins"
            Severity = "Medium"; Score = 40; MaxScore = 100
            EvidenceNote = "Synthetic governance gap"
        }
    )
}

Write-EvidenceJson -NamePrefix "Audit_Forwarding_SignIn_Assessment" -Object $findings | Out-Null

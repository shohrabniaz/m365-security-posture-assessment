<#
.SYNOPSIS
  Assess MFA registration posture (synthetic by default; -Live uses Graph when available).
#>
param([switch]$Live)
. "$PSScriptRoot\00_Config.ps1"

$useLive = Assert-LabTenantOrSynthetic -Live:$Live
$findings = [ordered]@{
    ControlFamily   = "MFA / authentication strength"
    Mode            = if ($useLive) { "LiveGraph" } else { "Synthetic" }
    Organisation    = $FictionalOrg
    CollectedUtc    = (Get-Date).ToUniversalTime().ToString("o")
    GuidanceRefs    = @(
        "Microsoft: Require MFA for all users (Secure Score / identity recommendations)",
        "ASD Essential Eight: Multi-factor authentication",
        "CIS Microsoft 365 Foundations Benchmark — MFA controls"
    )
}

if ($useLive) {
    # Read-only intent: report what Graph returns without changing tenant settings.
    $findings.Note = "Live MFA registration report requires Microsoft Graph Reports.Read.All or Identity.Read.All and appropriate licence surface. Capture results honestly; do not invent portal values."
    $findings.Items = @()
}
else {
    $findings.Items = @(
        [ordered]@{
            Id = "MFA-01"; Title = "MFA for privileged roles"
            Observed = "Global Admin and Privileged Role Admin modelled without enforced phishing-resistant MFA on Free tier"
            Severity = "High"; Score = 20; MaxScore = 100
            EvidenceNote = "Synthetic — Entra ID Free cannot claim Conditional Access MFA policies"
        },
        [ordered]@{
            Id = "MFA-02"; Title = "MFA for standard users"
            Observed = "Per-user MFA not consistently enabled for all modelled cloud users"
            Severity = "High"; Score = 35; MaxScore = 100
            EvidenceNote = "Synthetic baseline for CMIT ~60-user estate"
        },
        [ordered]@{
            Id = "MFA-03"; Title = "Legacy authentication"
            Observed = "Legacy auth protocols assumed still possible without CA block"
            Severity = "Medium"; Score = 40; MaxScore = 100
            EvidenceNote = "Synthetic — Free tier has no Conditional Access to block legacy auth"
        }
    )
}

Write-EvidenceJson -NamePrefix "MFA_Assessment" -Object $findings | Out-Null

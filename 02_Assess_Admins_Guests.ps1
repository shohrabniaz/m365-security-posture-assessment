<#
.SYNOPSIS
  Assess admin accounts and guest access posture.
#>
param([switch]$Live)
. "$PSScriptRoot\00_Config.ps1"

$useLive = Assert-LabTenantOrSynthetic -Live:$Live
$findings = [ordered]@{
    ControlFamily = "Admin accounts and guest access"
    Mode          = if ($useLive) { "LiveGraph" } else { "Synthetic" }
    Organisation  = $FictionalOrg
    CollectedUtc  = (Get-Date).ToUniversalTime().ToString("o")
    GuidanceRefs  = @(
        "Microsoft: Separate admin and daily accounts; minimise standing Global Admin",
        "ASD Essential Eight: Restrict administrative privileges",
        "Microsoft: Guest access reviews and external sharing governance"
    )
}

if ($useLive) {
    $findings.Note = "Live path: Get-MgDirectoryRole / Get-MgUser -Filter for userType eq Guest. Refuse if tenantId != lab."
    $findings.Items = @()
}
else {
    $findings.Items = @(
        [ordered]@{
            Id = "ADM-01"; Title = "Standing privileged group membership"
            Observed = "Modelled privileged security group includes a disabled orphan account (leaver residual) — pattern carried from prior JML lab narrative"
            Severity = "Critical"; Score = 10; MaxScore = 100
            EvidenceNote = "Synthetic / prior-lab pattern — DISABLED_BUT_STILL_IN_PRIV_GROUP class finding"
        },
        [ordered]@{
            Id = "ADM-02"; Title = "Number of Global Administrators"
            Observed = "More than two standing Global Admin equivalents modelled (break-glass + day-to-day overlap)"
            Severity = "High"; Score = 30; MaxScore = 100
            EvidenceNote = "Synthetic — recommend 2 break-glass + JIT for operational admin"
        },
        [ordered]@{
            Id = "GST-01"; Title = "Guest user inventory"
            Observed = "Guests present without periodic access review"
            Severity = "Medium"; Score = 45; MaxScore = 100
            EvidenceNote = "Synthetic — Free tier lacks Entra ID Governance access reviews"
        },
        [ordered]@{
            Id = "GST-02"; Title = "Guest invitation settings"
            Observed = "Members can invite guests (default-like posture)"
            Severity = "Medium"; Score = 50; MaxScore = 100
            EvidenceNote = "Synthetic — tighten to admins-only for SME MSP"
        }
    )
}

Write-EvidenceJson -NamePrefix "Admins_Guests_Assessment" -Object $findings | Out-Null

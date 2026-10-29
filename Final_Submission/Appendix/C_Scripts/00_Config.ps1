# Shared config for Academic Project 15 (laboratory / synthetic only)
$script:ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:EvidenceDir = Join-Path $ProjectRoot "evidence"
$script:LabTenantId = "a76a68a8-6243-4b0e-b11c-98e815e91820"
$script:LabTenantDomain = "niazkhanghostgmail.onmicrosoft.com"
$script:FictionalOrg = "Coastal Managed IT (CMIT) — fictional SME case"

if (-not (Test-Path $EvidenceDir)) {
    New-Item -ItemType Directory -Path $EvidenceDir | Out-Null
}

function Get-Stamp {
    Get-Date -Format "yyyyMMdd_HHmmss"
}

function Write-EvidenceJson {
    param(
        [Parameter(Mandatory = $true)][string]$NamePrefix,
        [Parameter(Mandatory = $true)]$Object
    )
    $path = Join-Path $EvidenceDir ("{0}_{1}.json" -f (Get-Stamp), $NamePrefix)
    $Object | ConvertTo-Json -Depth 8 | Set-Content -Path $path -Encoding UTF8
    Write-Host "Wrote $path"
    return $path
}

function Assert-LabTenantOrSynthetic {
    param([switch]$Live)
    if (-not $Live) { return $false }
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if (-not $ctx -or $ctx.TenantId -ne $LabTenantId) {
            throw "Signed-in tenant '$($ctx.TenantId)' is not the laboratory tenant $LabTenantId. Refusing live collection."
        }
        return $true
    }
    catch {
        Write-Warning "Live Graph unavailable ($($_.Exception.Message)). Falling back to synthetic baseline."
        return $false
    }
}

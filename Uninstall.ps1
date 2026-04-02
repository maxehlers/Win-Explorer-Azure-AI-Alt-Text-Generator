<#
.SYNOPSIS
    Removes the "Generate Alt-Text" context menu, deletes stored Azure credentials,
    and cleans up the Toast app registration.

.NOTES
    No administrator rights required (everything was installed under HKCU).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$CRED_FOUNDRY_ENDPOINT   = "AltTextGen:FoundryEndpoint"
$CRED_FOUNDRY_APIKEY     = "AltTextGen:FoundryApiKey"
$CRED_FOUNDRY_DEPLOYMENT = "AltTextGen:FoundryDeployment"
$CRED_LEGACY_ENDPOINT    = "AltTextGen:Endpoint"
$CRED_LEGACY_APIKEY      = "AltTextGen:ApiKey"
$APP_ID        = "AltTextGenerator.ContextMenu"

function Get-CredHelperType {
    param([string]$SourcePath)

    $hash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash.Substring(0, 10)
    $typeName = "WinCredHelper_$hash"

    $existing = ([System.Management.Automation.PSTypeName]$typeName).Type
    if ($existing) { return $existing }

    $src = Get-Content -LiteralPath $SourcePath -Raw
    $patched = $src -replace 'public\s+static\s+class\s+WinCredHelperV2', "public static class $typeName"
    Add-Type -TypeDefinition $patched -Language CSharp
    return ([System.Management.Automation.PSTypeName]$typeName).Type
}

Write-Host ""
Write-Host "  Explorer Alt Text Generator - Uninstall" -ForegroundColor Cyan
Write-Host "  ==========================================" -ForegroundColor Cyan
Write-Host ""

# ── Remove stored credentials ────────────────────────────────────────────────
$credSrc = Join-Path $PSScriptRoot "WinCredHelper.cs"
if (Test-Path $credSrc) {
    $CredHelperType = Get-CredHelperType -SourcePath $credSrc

    foreach ($target in @(
            $CRED_FOUNDRY_ENDPOINT,
            $CRED_FOUNDRY_APIKEY,
            $CRED_FOUNDRY_DEPLOYMENT,
            $CRED_LEGACY_ENDPOINT,
            $CRED_LEGACY_APIKEY)) {
        if ($CredHelperType::Exists($target)) {
            $CredHelperType::Delete($target)
            Write-Host "  [OK] Removed credential: $target" -ForegroundColor Green
        }
    }
}
else {
    Write-Host "  [--] WinCredHelper.cs not found; skipping credential removal." -ForegroundColor Yellow
}

# ── Remove context menu registry keys ───────────────────────────────────────
$shellKey = "HKCU:\Software\Classes\SystemFileAssociations\image\shell\GenerateAltText"
if (Test-Path $shellKey) {
    Remove-Item -Path $shellKey -Recurse -Force
    Write-Host "  [OK] Context menu entry removed." -ForegroundColor Green
}
else {
    Write-Host "  [--] Context menu entry not found (already removed?)." -ForegroundColor Yellow
}

# ── Remove Toast app registration ───────────────────────────────────────────
$toastKey = "HKCU:\Software\Classes\AppUserModelId\$APP_ID"
if (Test-Path $toastKey) {
    Remove-Item -Path $toastKey -Recurse -Force
    Write-Host "  [OK] Toast notification registration removed." -ForegroundColor Green
}

Write-Host ""
Write-Host "  Uninstall complete." -ForegroundColor Cyan
Write-Host "  You can safely delete the script folder afterwards." -ForegroundColor Cyan
Write-Host ""

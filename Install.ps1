<#
.SYNOPSIS
    Installs the "Generate Alt-Text" context menu entry for image files in Windows Explorer.

.DESCRIPTION
    - Prompts for your Azure AI Foundry / Azure OpenAI endpoint, API key, and deployment name
    - Stores credentials securely in Windows Credential Manager (no plain-text files)
    - Registers a right-click context menu entry for all image files via HKCU registry
    - No administrator rights required

.NOTES
    Prerequisites : Azure AI Foundry project endpoint OR Azure OpenAI endpoint,
                    plus a deployed multimodal model (e.g. gpt-4o / gpt-4.1).

    Run again at any time to update credentials or re-register the menu entry.
    Compatible with: Windows 10 / 11, Windows PowerShell 5.1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Constants ────────────────────────────────────────────────────────────────
$CRED_FOUNDRY_ENDPOINT   = "AltTextGen:FoundryEndpoint"
$CRED_FOUNDRY_APIKEY     = "AltTextGen:FoundryApiKey"
$CRED_FOUNDRY_DEPLOYMENT = "AltTextGen:FoundryDeployment"
$APP_ID        = "AltTextGenerator.ContextMenu"
$MENU_LABEL    = "Generate Alt-Text"
$ICON_SVG      = Join-Path $PSScriptRoot "favicon.svg"
$ICON_PNG      = Join-Path $PSScriptRoot "favicon.png"
$ICON_ICO      = Join-Path $PSScriptRoot "favicon.ico"

# ── Sanity checks ────────────────────────────────────────────────────────────
$credSrc    = Join-Path $PSScriptRoot "WinCredHelper.cs"
$mainScript = Join-Path $PSScriptRoot "GenerateAltText.ps1"

foreach ($f in @($credSrc, $mainScript)) {
    if (-not (Test-Path $f)) {
        Write-Error "Required file not found: $f"
        exit 1
    }
}

# ── Load credential helper ───────────────────────────────────────────────────
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

$CredHelperType = Get-CredHelperType -SourcePath $credSrc

# ── Banner ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Explorer Alt Text Generator - Setup" -ForegroundColor Cyan
Write-Host "  =====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Requires Azure AI Foundry or Azure OpenAI with a deployed vision-capable model."
Write-Host "  Examples:"
Write-Host "    Endpoint (Azure OpenAI): https://<resource>.openai.azure.com"
Write-Host "    Endpoint (AI Foundry):   https://<project>.services.ai.azure.com"
Write-Host "    Deployment name:         e.g. gpt-4o"
Write-Host ""

# ── Collect credentials ──────────────────────────────────────────────────────
$saveCreds = $true
if ($CredHelperType::Exists($CRED_FOUNDRY_ENDPOINT)) {
    $answer    = Read-Host "  Credentials already saved. Update them? [y/N]"
    $saveCreds = ($answer -match "^[yY]")
}

if ($saveCreds) {
    $endpoint = ""
    while ([string]::IsNullOrWhiteSpace($endpoint)) {
        $endpoint = (Read-Host "  Azure AI Foundry / Azure OpenAI Endpoint`n  (e.g. https://myresource.openai.azure.com)").Trim().TrimEnd("/")
    }

    $deployment = ""
    while ([string]::IsNullOrWhiteSpace($deployment)) {
        $deployment = (Read-Host "  Deployment name`n  (e.g. gpt-4o)").Trim()
    }

    $secureKey = Read-Host "  API Key" -AsSecureString
    $apiKey    = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                     [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey))

    $CredHelperType::SaveSecret($CRED_FOUNDRY_ENDPOINT,   $endpoint)
    $CredHelperType::SaveSecret($CRED_FOUNDRY_APIKEY,     $apiKey)
    $CredHelperType::SaveSecret($CRED_FOUNDRY_DEPLOYMENT, $deployment)

    Write-Host ""
    Write-Host "  [OK] Credentials saved to Windows Credential Manager." -ForegroundColor Green
}

# ── Register Toast App ID (needed for modern toast notifications) ─────────────
$toastRegPath = "HKCU:\Software\Classes\AppUserModelId\$APP_ID"
New-Item -Path $toastRegPath -Force | Out-Null
Set-ItemProperty -Path $toastRegPath -Name "DisplayName" -Value "Alt Text Generator"
Write-Host "  [OK] Toast notification app registered." -ForegroundColor Green

# ── Register context menu (HKCU – no admin rights required) ──────────────────
#    Applies to all image files via SystemFileAssociations\image
$pwshExe  = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$cmdValue = "`"$pwshExe`" -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$mainScript`" `"%1`""

$menuIconValue = "imageres.dll,-5302"
if (Test-Path -LiteralPath $ICON_ICO) {
    $menuIconValue = "$ICON_ICO,0"
}
elseif (Test-Path -LiteralPath $ICON_PNG) {
    $menuIconValue = "$ICON_PNG,0"
}
elseif (Test-Path -LiteralPath $ICON_SVG) {
    $magick = Get-Command magick -ErrorAction SilentlyContinue
    if ($magick) {
        try {
            & magick "$ICON_SVG" -background none -resize 256x256 "$ICON_ICO"
            if (Test-Path -LiteralPath $ICON_ICO) {
                $menuIconValue = "$ICON_ICO,0"
                Write-Host "  [OK] favicon.svg was converted to favicon.ico." -ForegroundColor Green
            }

            if (-not (Test-Path -LiteralPath $ICON_PNG)) {
                & magick "$ICON_SVG" -background none -resize 256x256 "$ICON_PNG"
                if (Test-Path -LiteralPath $ICON_PNG) {
                    Write-Host "  [OK] favicon.svg was converted to favicon.png for toasts." -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host "  [--] SVG conversion failed. Using default Explorer icon." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  [--] Found favicon.svg but no ImageMagick (magick) to convert to .ico. Using default Explorer icon." -ForegroundColor Yellow
    }
}

$shellKey = "HKCU:\Software\Classes\SystemFileAssociations\image\shell\GenerateAltText"
New-Item -Path $shellKey           -Force | Out-Null
New-Item -Path "$shellKey\command" -Force | Out-Null

Set-ItemProperty -Path $shellKey           -Name "(Default)" -Value $MENU_LABEL
Set-ItemProperty -Path $shellKey           -Name "Icon"      -Value $menuIconValue
Set-ItemProperty -Path "$shellKey\command" -Name "(Default)" -Value $cmdValue

Write-Host "  [OK] Context menu registered for all image files." -ForegroundColor Green

# ── Done ─────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Installation complete!" -ForegroundColor Cyan
Write-Host "  Right-click any image in Explorer and select '$MENU_LABEL'." -ForegroundColor Cyan
Write-Host ""

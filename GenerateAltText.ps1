<#
.SYNOPSIS
    Called by the Windows Explorer context menu to generate alt text for an image.

.DESCRIPTION
    Reads Azure AI Foundry / Azure OpenAI credentials from Windows Credential Manager,
    sends the image to a multimodal model, copies the generated alt text to the
    clipboard, and shows a Windows toast notification.

.PARAMETER ImagePath
    Full path to the image file as passed by the Explorer context menu (%1).
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath
)

$ErrorActionPreference = "Stop"

# Constants
$CRED_FOUNDRY_ENDPOINT   = "AltTextGen:FoundryEndpoint"
$CRED_FOUNDRY_APIKEY     = "AltTextGen:FoundryApiKey"
$CRED_FOUNDRY_DEPLOYMENT = "AltTextGen:FoundryDeployment"
$APP_ID                  = "AltTextGenerator.ContextMenu"
$MAX_FILE_SIZE           = 4MB

function Get-ToastLogoPath {
    $candidates = @(
        (Join-Path $PSScriptRoot "toast-logo.png"),
        (Join-Path $PSScriptRoot "favicon.png"),
        (Join-Path $PSScriptRoot "favicon.ico"),
        (Join-Path $PSScriptRoot "favicon.svg")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Initialize-CredHelper {
    $src = Join-Path $PSScriptRoot "WinCredHelper.cs"
    if (-not (Test-Path $src)) {
        throw "WinCredHelper.cs not found. Please re-run Install.ps1."
    }

    $hash = (Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash.Substring(0, 10)
    $typeName = "WinCredHelper_$hash"

    $existing = ([System.Management.Automation.PSTypeName]$typeName).Type
    if ($existing) {
        return $existing
    }

    $code = Get-Content -LiteralPath $src -Raw
    $patched = $code -replace 'public\s+static\s+class\s+WinCredHelperV2', "public static class $typeName"
    Add-Type -TypeDefinition $patched -Language CSharp
    return ([System.Management.Automation.PSTypeName]$typeName).Type
}

function Show-Toast {
    param(
        [string]$Title,
        [string]$Body
    )
    try {
        $t   = [System.Security.SecurityElement]::Escape($Title)
        $b   = [System.Security.SecurityElement]::Escape($Body)
        $logoNode = ""
        $logoPath = Get-ToastLogoPath
        if (-not [string]::IsNullOrWhiteSpace($logoPath)) {
            $logoUri = ([System.Uri]::new((Resolve-Path -LiteralPath $logoPath))).AbsoluteUri
            $logoNode = '<image placement="appLogoOverride" hint-crop="circle" src="' + $logoUri + '"/>'
        }

        $xml = '<toast><visual><binding template="ToastGeneric">' +
               $logoNode +
               '<text>' + $t + '</text><text>' + $b + '</text></binding></visual></toast>'

        # Keep WinRT type syntax inside dynamic script to avoid editor parser false positives.
        $sb = [scriptblock]::Create(@'
            param($xmlString, $appId)
            [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
            [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]
            $doc = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]::new()
            $doc.LoadXml($xmlString)
            $notifier = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::CreateToastNotifier($appId)
            $toast    = [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime]::new($doc)
            $notifier.Show($toast)
'@)
        & $sb $xml $APP_ID
    }
    catch {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            $icon         = [System.Windows.Forms.NotifyIcon]::new()
            $icon.Icon    = [System.Drawing.SystemIcons]::Information
            $icon.Visible = $true
            $icon.ShowBalloonTip(5000, $Title, $Body, [System.Windows.Forms.ToolTipIcon]::Info)
            Start-Sleep -Seconds 6
            $icon.Dispose()
        }
        catch { }
    }
}

function Show-Error([string]$Message) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            "Alt Text Generator",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    catch { }
    Show-Toast -Title "Alt Text Generator - Error" -Body $Message
}

function Get-ImageMimeType {
    param([string]$Path)

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($ext) {
        ".jpg"  { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".png"  { return "image/png" }
        ".gif"  { return "image/gif" }
        ".bmp"  { return "image/bmp" }
        ".webp" { return "image/webp" }
        default  { return "image/jpeg" }
    }
}

function Get-ImageBytesForApi {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [int64]$MaxBytes
    )

    $originalBytes = [System.IO.File]::ReadAllBytes($Path)
    if ($originalBytes.Length -le $MaxBytes) {
        return @{
            Bytes         = $originalBytes
            MimeType      = Get-ImageMimeType -Path $Path
            WasCompressed = $false
        }
    }

    Add-Type -AssemblyName System.Drawing

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $sourceImage = [System.Drawing.Image]::FromStream($stream, $true, $true)
    }
    finally {
        $stream.Dispose()
    }

    try {
        $jpegCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
            Where-Object { $_.MimeType -eq "image/jpeg" } |
            Select-Object -First 1

        if (-not $jpegCodec) {
            throw "JPEG encoder not available on this system."
        }

        $qualitySteps = @(90L, 80L, 70L, 60L, 50L, 40L, 30L, 20L)
        $scale        = 1.0

        for ($resizeAttempt = 0; $resizeAttempt -lt 4; $resizeAttempt++) {
            $targetWidth  = [Math]::Max([int]($sourceImage.Width * $scale), 1)
            $targetHeight = [Math]::Max([int]($sourceImage.Height * $scale), 1)

            $bitmap = [System.Drawing.Bitmap]::new($targetWidth, $targetHeight)
            try {
                $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                try {
                    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                    $graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                    $graphics.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                    $graphics.DrawImage($sourceImage, 0, 0, $targetWidth, $targetHeight)
                }
                finally {
                    $graphics.Dispose()
                }

                foreach ($quality in $qualitySteps) {
                    $encoderParams = [System.Drawing.Imaging.EncoderParameters]::new(1)
                    $encoderParams.Param[0] = [System.Drawing.Imaging.EncoderParameter]::new(
                        [System.Drawing.Imaging.Encoder]::Quality,
                        $quality)

                    $memory = [System.IO.MemoryStream]::new()
                    try {
                        $bitmap.Save($memory, $jpegCodec, $encoderParams)
                        if ($memory.Length -le $MaxBytes) {
                            return @{
                                Bytes         = $memory.ToArray()
                                MimeType      = "image/jpeg"
                                WasCompressed = $true
                            }
                        }
                    }
                    finally {
                        $memory.Dispose()
                        $encoderParams.Dispose()
                    }
                }
            }
            finally {
                $bitmap.Dispose()
            }

            $scale *= 0.75
        }
    }
    finally {
        $sourceImage.Dispose()
    }

    throw "Image is too large and could not be compressed below 4 MB."
}

function Get-ChatTextFromResponse {
    param([object]$Response)

    $content = $null

    if ($Response -and $Response.choices -and $Response.choices.Count -gt 0) {
        $content = $Response.choices[0].message.content
    }

    if ([string]::IsNullOrWhiteSpace([string]$content) -eq $false) {
        return ([string]$content).Trim()
    }

    if ($content -is [System.Collections.IEnumerable]) {
        $parts = @()
        foreach ($item in $content) {
            if ($item -and $item.text) {
                $parts += [string]$item.text
            }
        }
        $joined = ($parts -join " ").Trim()
        if (-not [string]::IsNullOrWhiteSpace($joined)) {
            return $joined
        }
    }

    return $null
}

function Invoke-FoundryAltText {
    param(
        [string]$Endpoint,
        [string]$ApiKey,
        [string]$Deployment,
        [string]$DataUrl
    )

    $endpoint = $Endpoint.Trim().TrimEnd("/")
    $systemPrompt = "Du erzeugst hochwertigen Alt-Text fuer Screenreader. Gib genau einen klaren Satz aus, keine Einleitung, keine Aufzaehlung. Nenne die wichtigsten sichtbaren Inhalte und lesbaren Kerntext (z. B. Eventtitel, Datum, Uhrzeit, Ort), sofern vorhanden. Bitte gib den Alt-Text in der ursprunglichen Sprache des Textes im Bild zurück und wenn Du diese nicht erkennst, bitte auf Englisch. Wenn das Bild keinen lesbaren Text oder erkennbare Inhalte hat, gib eine kurze allgemeine Beschreibung zurück (z. B. 'Photo of a person', 'screenshot of a chat conversation'). Antworte nur mit dem Alt-Text, ohne weitere Erklärungen oder Anmerkungen."
    $userPrompt = "Erzeuge einen praezisen Alt-Text fuer dieses Bild."

    $contentParts = @(
        @{ type = "text"; text = $userPrompt },
        @{ type = "image_url"; image_url = @{ url = $DataUrl; detail = "high" } }
    )

    $attempts = New-Object System.Collections.Generic.List[hashtable]

    if ($endpoint -like "*.openai.azure.com") {
        if ([string]::IsNullOrWhiteSpace($Deployment)) {
            throw "Missing deployment name for Azure OpenAI endpoint."
        }

        foreach ($apiVersion in @("2024-10-21", "2024-06-01")) {
            $uri = "$endpoint/openai/deployments/$Deployment/chat/completions?api-version=$apiVersion"
            $body = @{
                messages = @(
                    @{ role = "system"; content = $systemPrompt },
                    @{ role = "user"; content = $contentParts }
                )
                max_tokens  = 3000
                temperature = 0.2
            }
            $attempts.Add(@{ Uri = $uri; Body = $body; Headers = @{ "api-key" = $ApiKey } })
        }
    }

    # AI Foundry / Azure AI Inference compatible attempts.
    foreach ($path in @("/models/chat/completions", "/chat/completions")) {
        foreach ($apiVersion in @("2024-05-01-preview", "2024-02-15-preview")) {
            $uri = "$endpoint$path?api-version=$apiVersion"

            $bodyWithModel = @{
                messages = @(
                    @{ role = "system"; content = $systemPrompt },
                    @{ role = "user"; content = $contentParts }
                )
                max_tokens  = 220
                temperature = 0.2
                model       = $Deployment
            }

            $bodyWithoutModel = @{
                messages = @(
                    @{ role = "system"; content = $systemPrompt },
                    @{ role = "user"; content = $contentParts }
                )
                max_tokens  = 220
                temperature = 0.2
            }

            $attempts.Add(@{ Uri = $uri; Body = $bodyWithModel; Headers = @{ "api-key" = $ApiKey } })
            $attempts.Add(@{ Uri = $uri; Body = $bodyWithoutModel; Headers = @{ "api-key" = $ApiKey } })
        }
    }

    $lastError = $null
    foreach ($attempt in $attempts) {
        try {
            $json = $attempt.Body | ConvertTo-Json -Depth 30
            $resp = Invoke-RestMethod -Uri $attempt.Uri -Method POST -Headers $attempt.Headers -ContentType "application/json" -Body $json
            $text = Get-ChatTextFromResponse -Response $resp
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                return $text
            }
        }
        catch {
            $lastError = $_
        }
    }

    if ($lastError) {
        throw "Foundry call failed. Last error: $($lastError.Exception.Message)"
    }

    throw "Foundry call failed. No response text returned."
}

function Normalize-AltText {
    param([string]$Text)

    $normalized = [regex]::Replace($Text, "\s+", " ").Trim()
    $normalized = $normalized.Trim('"')

    return $normalized
}

try {
    $CredHelperType = Initialize-CredHelper

    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "File not found:`n$ImagePath"
    }

    $file = Get-Item -LiteralPath $ImagePath

    $endpoint   = $CredHelperType::GetSecret($CRED_FOUNDRY_ENDPOINT)
    $apiKey     = $CredHelperType::GetSecret($CRED_FOUNDRY_APIKEY)
    $deployment = $CredHelperType::GetSecret($CRED_FOUNDRY_DEPLOYMENT)

    if ([string]::IsNullOrWhiteSpace($endpoint) -or
        [string]::IsNullOrWhiteSpace($apiKey) -or
        [string]::IsNullOrWhiteSpace($deployment)) {
        throw "Foundry credentials are not configured.`nPlease run Install.ps1 again."
    }

    $payload = Get-ImageBytesForApi -Path $ImagePath -MaxBytes $MAX_FILE_SIZE
    $base64  = [System.Convert]::ToBase64String($payload.Bytes)
    $dataUrl = "data:$($payload.MimeType);base64,$base64"

    Show-Toast -Title "Alt-Text wird erstellt" -Body "Bild wird an die KI gesendet..."

    $altText = Invoke-FoundryAltText -Endpoint $endpoint -ApiKey $apiKey -Deployment $deployment -DataUrl $dataUrl
    $altText = Normalize-AltText -Text $altText

    if ([string]::IsNullOrWhiteSpace($altText)) {
        throw "No usable alt text was returned by the model."
    }

    try {
        Set-Clipboard -Value $altText
    }
    catch {
        throw "Alt text was generated but could not be copied to clipboard.`n`nAlt text: $altText`n`nError: $($_.Exception.Message)"
    }

    Show-Toast -Title "Alt text copied" -Body "`"$altText`"`n($($file.Name))"
}
catch {
    Show-Error "Alt text could not be generated.`n`n$($_.Exception.Message)"
}

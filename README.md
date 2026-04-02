# Explorer Alt Text Generator 🖼️✨

Create high-quality alt text from Windows Explorer with one right-click.

This tool adds a **Generate Alt-Text** context menu entry for image files. It sends the image to a multimodal model in **Azure AI Foundry / Azure OpenAI**, copies the generated alt text to your clipboard, and shows Windows toast notifications.

## What You Get 🚀

- Right-click image -> **Generate Alt-Text**
- AI-generated, richer alt text (LLM + vision)
- Auto-copy to clipboard 📋
- Progress + success/error toast notifications 🔔
- Secure credential storage in Windows Credential Manager 🔐
- Automatic image compression when file size is too large 🗜️

## Prerequisites ✅

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- Azure AI Foundry or Azure OpenAI endpoint
- A deployed multimodal model (for example: `gpt-4o`)

Typical endpoints:

- Azure OpenAI: `https://<resource>.openai.azure.com`
- Azure AI Foundry project endpoint: `https://<project>.services.ai.azure.com`

## Step-by-Step Installation 🛠️

### 1. Put all files in one folder 📁

Make sure these files stay together:

- `Install.ps1`
- `GenerateAltText.ps1`
- `Uninstall.ps1`
- `WinCredHelper.cs`
- `favicon.svg` (optional but recommended for custom icon)

Do not move files after installation unless you run install again.

### 2. Open PowerShell in this folder 💻

Example:

```powershell
cd "C:\Users\<you>\...\Explorer Alt Text Generator"
```

### 3. Run the installer ▶️

If scripts are allowed:

```powershell
.\Install.ps1
```

If blocked by execution policy:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Install.ps1"
```

### 4. Enter credentials when prompted 🔑

The installer asks for:

1. Endpoint
2. Deployment name (example: `gpt-4o`)
3. API key

Credentials are stored in **Windows Credential Manager** (not in plain text files).

### 5. Optional: custom icon setup 🎨

- If `favicon.ico` exists, Explorer menu uses it.
- If only `favicon.svg` exists and ImageMagick (`magick`) is installed, install tries to auto-convert to `favicon.ico` and `favicon.png`.
- Toast logo prefers, in order:
  - `toast-logo.png`
  - `favicon.png`
  - `favicon.ico`
  - `favicon.svg`

## Step-by-Step Usage 🧭

1. Open Windows Explorer.
2. Right-click an image file (`.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.webp`, ...).
3. Click **Generate Alt-Text**.
4. You will see a progress toast: image is being sent to AI.
5. After completion, a success toast appears.
6. Paste anywhere with `Ctrl + V`.

## How It Works Internally ⚙️

1. Validates and optionally compresses the image (target <= 4 MB).
2. Encodes image as data URL.
3. Calls Azure AI Foundry / Azure OpenAI chat-completions endpoint.
4. Extracts and normalizes full model output (no forced truncation).
5. Copies text to clipboard.
6. Shows toast notifications.

## Updating Credentials 🔄

Run installer again:

```powershell
.\Install.ps1
```

It asks whether you want to overwrite saved credentials.

## Uninstall 🧹

```powershell
.\Uninstall.ps1
```

This removes:

- Explorer context menu entry
- Toast app registration
- Stored credentials (Foundry/OpenAI + legacy keys)

## Troubleshooting 🩺

### Script execution blocked

Use:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Install.ps1"
```

### 401 Unauthorized

- API key is wrong, expired, or for another resource.
- Re-run install and save correct key.

### 404 Not Found

- Endpoint or deployment name is wrong.
- Verify endpoint URL and deployment name in Azure.

### No menu icon update

- Re-run install.
- Restart Explorer if Windows icon cache delays updates.

### Toast not visible

- Check Windows notification settings.
- Fallback balloon notification is used when WinRT toast is unavailable.

## Security Notes 🔐

- API credentials are stored in Windows Credential Manager.
- No credentials are hardcoded in scripts.
- No plaintext secrets are written to project files.

## File Overview 📚

| File | Purpose |
|---|---|
| `Install.ps1` | Setup wizard, credential capture, context menu registration |
| `GenerateAltText.ps1` | Main worker script (API call, clipboard, toasts) |
| `Uninstall.ps1` | Cleanup script (menu + credentials + toast registration) |
| `WinCredHelper.cs` | Native interop helper for Credential Manager |
| `favicon.svg` | Source icon asset for menu/toast branding |
